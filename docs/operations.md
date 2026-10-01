# Operations and troubleshooting

## Customer operation

### Add or change an agent

Publish and share the agent, then add or edit its SharePoint row. No Teams
rebuild is needed.

### Disable an agent

Set `Enabled` to false. Existing browser sessions can retain their loaded list
until reconnect; new connections omit the agent.

### Change voice or avatar

Edit `Locale`, `VoiceName`, `AvatarCharacter`, and `AvatarStyle`. Selecting the
agent recreates the avatar session when the profile changed.

## Health checks

```powershell
Invoke-RestMethod https://<app>.azurewebsites.net/api/health
Invoke-RestMethod https://<app>.azurewebsites.net/api/config
```

The config response contains public app/catalog defaults only. It must never
contain the Speech key.

## Authentication troubleshooting

### Authentication window shows another copy of the app

The Entra SPA callback is wrong or stale. It must be:

```text
https://<app>.azurewebsites.net/auth/callback
```

Re-run provisioning after updating `aad.manifest.json`.

### More than one first-use consent window

Expected when Graph and Power Platform have not been pre-consented. They are
different token resources. An administrator can pre-consent both delegated
permissions.

### Wrong account

Sign out of the relevant cache (Azure CLI, Agents Toolkit Azure, Agents Toolkit
Microsoft 365, or browser) and retry. Always state the required customer account
before starting authentication.

## Catalog troubleshooting

- Confirm the user has SharePoint read access.
- Confirm Graph consent.
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
