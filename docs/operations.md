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

The config response contains public app/catalog defaults and the
`api://<client-id>/access_as_user` scope only. It must never contain the Speech
key.

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

Standalone consent now uses same-window redirect. The callback bridge returns
to the app, processes the result and reloads the authorized catalog. A redirect
does not start an agent or repeat a business action. Teams uses its nested broker
instead of redirecting its tab.

### Need admin approval

The generic app requests its own `access_as_user` scope and
`CopilotStudio.Copilots.Invoke`; it requests no Graph permission. If tenant
policy blocks user consent, ask an administrator to grant tenant-wide consent
for these delegated permissions after provisioning the updated app manifest.
SharePoint catalog access uses the maker-owned flow and does not request Graph
consent from app users.

### AADSTS65006 or an invalid invocation permission

Resolve **CopilotStudio.Copilots.Invoke** by name from the tenant's Power
Platform API service principal before declaring the permission. Do not substitute
a similar maker/admin scope or grant broad permissions. A stale permission GUID
can make `.default` token acquisition fail even when the agent is published.
Update only this app's required resource access, then retry user consent.

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
  provision again, and keep production connection unavailable. Do not use a
  fallback agent to bypass catalog governance.
- POST the test request in `catalog/power-automate-flow.md`.
- Confirm required display-name columns exist.
- Do not assume CSV-imported internal field names; the app resolves columns by
  display name in the flow field map.
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

### Avatar lifecycle and cost control

Text-only is the default. **Read typed replies aloud** and **Avatar video** are
explicit paid-mode choices; **Start voice conversation** opts into recognition
and spoken replies. Estimated voice and avatar rates appear before the controls.

Avatar synthesis is permitted only after a successful, non-ended Copilot Studio
connection. Demo mode does not qualify. Each avatar reply closes its synthesizer,
WebRTC connection and received tracks when it finishes, so there is no paid
avatar waiting between turns. Completion first releases the avatar; any final
spoken confirmation uses audio-only playback.

Stop, avatar-off, agent changes, errors, page exit and hidden tabs release
resources. Late credentials, startup results and media tracks cannot revive a
released session. A hung avatar start is capped at 30 seconds; a hung avatar
reply at 120 seconds. Media failures also trigger cleanup. Azure's service-side
idle timeout is a fallback for browser crashes, not the normal cleanup path.
There is no central browser-session enumeration or server-side orphan sweeper.

The rule uses the client's confirmed conversation and observed completion/errors;
the SDK does not provide a continuous server-session liveness heartbeat. Remote
expiry is detected on the next request; no avatar stays connected between requests.

### Change cost wording without redeploying Teams

Set the optional App Service environment variables `VOICE_COST_NOTICE` and
`AVATAR_COST_NOTICE`. The backend returns these plain-text notices in
`/api/config`; no Teams package or code rebuild is needed. Saving App Service
settings restarts the host; reopen/refresh the app afterward. Never put credentials
or internal billing details in these public notices. Empty values use the
documented USD reference estimates, not measured bills.

For example, use **Azure portal > App Service > Settings > Environment variables**.
Update the notices when rates, region, currency or your contract change.
Per-agent names, descriptions and welcome messages remain SharePoint-managed.
Fixed interface labels still require a web-app build/deployment, not Teams approval.

Audio-only mode waits for the browser audio element's playback-end event before reopening
recognition, not merely the service's synthesis-completed callback. Stop pauses
the speaker and releases that wait. Audio-only output explicitly uses RIFF WAV
with native browser playback; the SDK's streaming player can stall without
delivering playback-end. Temporary audio URLs are revoked on completion or Stop.
Stop also invalidates pending Speech credential requests so late responses cannot
reopen the microphone or start playback.
CSP permits the Speech SDK's data/blob
workers separately; page scripts remain restricted to the same origin.

## Deployment troubleshooting

- Zero App Service quota: choose another avatar-supported region and revalidate.
- Key Vault policy requires purge protection: keep purge protection enabled.
- Teams package invalid: run Agents Toolkit package validation and inspect the
  generated manifest, not only the template.
- ZIP deployment is slow: wait for the active deployment; do not start
  overlapping deployments.
- Windows dependency restore reports `EBUSY`: do not kill unrelated VS Code or
  system processes. Run `scripts\stage-release.ps1` and deploy its isolated,
  ignored directory through the same lifecycle. This performs a clean locked
  restore without deleting the active workspace's dependencies.

## Cost and cleanup

Use [the release checklist](release-validation.md) before catalog submission.
Automatic URL/body/console tracing is disabled by default; aggregate metrics
are retained in Log Analytics for 30 days. Assign an operational and privacy
owner before wider rollout.

App Service B1 is the predictable fixed charge. Speech/avatar and telemetry are
usage-based. Use a dedicated resource group and budget alerts.

Resource-group deletion is destructive and must be explicitly approved. A
purge-protected Key Vault remains recoverable during retention.
