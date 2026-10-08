# Operations and troubleshooting

## Customer operation

### Add or change an agent

Publish and share the agent, then add or edit its SharePoint row. No Teams
rebuild is needed.

### Disable an agent

Set `Enabled` to false. Existing browser sessions can retain their loaded list
until reconnect; new connections omit the agent.

### Change voice or avatar

Edit output `Locale`, `VoiceName`, `AvatarCharacter`, and `AvatarStyle`. The
user chooses the STT input language separately in the app. Selecting the
agent recreates the avatar session when the profile changed.

## Health checks

```powershell
Invoke-RestMethod https://<app>.azurewebsites.net/api/health
Invoke-RestMethod https://<app>.azurewebsites.net/api/config
```

The config response contains public app/catalog defaults only. It must never
contain the Speech key.

## Authentication troubleshooting

### Teams asks for a separate account login

Teams should use nested app authentication instead. Confirm the Entra SPA has
`brk-multihub://<app-domain>`, re-run provisioning, and use a current Teams
desktop or web client. Older clients that do not support NAA are rejected rather
than opening a nested login popup.

### Authentication window shows another copy of the app

This browser-diagnostics fallback indicates that the Entra SPA callback is
wrong or stale. It must be:

```text
https://<app>.azurewebsites.net/auth/callback
```

Re-run provisioning after updating `aad.manifest.json`.

### Need admin approval

The generic app requests only `CopilotStudio.Copilots.Invoke`. If tenant policy
blocks user consent, ask an administrator to consent that delegated permission.
SharePoint catalog access uses the maker-owned flow and does not request Graph
consent from app users.

If user consent is allowed, accept the one-time Copilot Studio permission
dialog. Choosing **Return to app without granting permission** cancels the
connection; it cannot authorize the agent.

### Wrong account

Sign out of the relevant cache (Azure CLI, Agents Toolkit Azure, Agents Toolkit
Microsoft 365, or browser) and retry. Always state the required customer account
before starting authentication.

## Catalog troubleshooting

- Confirm the flow maker's SharePoint connection is healthy.
- Confirm `CATALOG_FLOW_URL` resolves from Key Vault.
- If the flow is intentionally not ready, set `CATALOG_FLOW_URL=disabled`,
  provision again, and use the configured default agent.
- POST the test request in `catalog/power-automate-flow.md`.
- Confirm required display-name columns exist.
- Do not assume CSV-imported internal field names; the app resolves columns by
  display name.
- Confirm environment IDs are GUIDs and schema names are published schema names.

## Speech troubleshooting

If `/api/speech/token` returns 502:

1. Check the Speech resource is S0 and succeeded.
2. Check the Key Vault secret exists.
3. Check the App Service identity has secret `get`.
4. Check App Service Key Vault reference status.
5. Refresh the reference and restart the app if it was authorized after initial
   app creation.

Avatar relay failure falls back to audio-only TTS. Basic TTS/STT can still work
when real-time avatar relay is unavailable.

## Deployment troubleshooting

- Zero App Service quota: choose another avatar-supported region and revalidate.
- Key Vault policy requires purge protection: keep purge protection enabled.
- Teams package invalid: run Agents Toolkit package validation and inspect the
  generated manifest, not only the template.
- ZIP deployment is slow: wait for the active deployment; do not start
  overlapping deployments.

## Cost and cleanup

App Service B1 is the predictable fixed charge. Speech/avatar and telemetry are
usage-based. Use a dedicated resource group and budget alerts.

Resource-group deletion is destructive and must be explicitly approved. A
purge-protected Key Vault remains recoverable during retention.
