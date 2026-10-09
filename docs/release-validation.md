# Governed catalog release gates

This checklist describes release requirements, not a claim that a deployment
has passed them. Keep customer-specific evidence in ignored `.azure/` files.
Do not submit the Teams package until each required gate is evidenced.

| Gate | Required evidence |
| --- | --- |
| Identity | Correct tenant/account; app-owned API scope; separate delegated Copilot permission; correct NAA and browser redirects |
| Protected APIs | Anonymous, invalid, expired, wrong-tenant, wrong-audience, and missing-scope requests rejected before catalog/Speech access |
| Catalog | Live maker-owned flow; valid enabled rows; disabled rows omitted; unavailable catalog stops connection rather than selecting a fallback |
| Audience | Everyone, allowed Restricted, and denied Restricted verified; caller comes from validated claims; browser response contains no AllowedUsers |
| Maintenance | Only list managers/site administrators receive the HTTPS configuration URL |
| Harness | Standard invocation verified; unsupported harness rows labelled and non-invokable |
| Voice | Selected STT input language; catalog output language/voice; avatar startup or explicitly identified audio-only failure |
| Rich content | Images, file links, citations, cards and structured actions render; attachment bodies are not narrated; unsafe URLs/actions rejected |
| Workflow | Pat follows up on insufficient input, sends exactly one HTML email, and ends only after confirmed success |
| Channels | Published-agent conversation verified in standalone browser and Teams; a maker test-panel result alone is insufficient |
| Privacy | No transcript, generated HTML, token, signed flow URL or people-picker data in telemetry; flow run-history inputs/outputs protected |
| Engineering | Tests, typecheck, production build, Bicep compilation, manifest validation and diff check pass |
| Operations | Monitoring ownership, consent/sharing, retention, incident response, rollback and cost ownership recorded |
| Release | Verified source commit and package version match deployment; tenant catalog approval recorded separately from submission |

## Automated checks

```powershell
npm ci
npm test
npm run typecheck
npm run build
az bicep build --file infra\azure.bicep
git diff --check
```

GitHub Actions and the deployment lifecycle run tests/typecheck/build.
Cloud what-if, consent, published-agent access, microphone/Teams behavior and
email delivery remain environment-specific acceptance checks, not unit tests.

## Voice acceptance evidence

After loading the catalog, **Start voice conversation** connects the selected
agent and starts the managed voice loop without a separate Connect action.
Verify a recognized answer is submitted after silence, recognition pauses for
the agent reply, listening resumes, and **Stop conversation** closes the audio
input. Repeat after restarting voice without reconnecting the agent.

Automation may inject nonsensitive synthetic audio instead of opening a
physical microphone. Record that distinction explicitly: synthetic tests prove
the recognition and turn-management path, not a user's headset, device
permissions, acoustic conditions or Teams desktop microphone behavior. Pilot
users must check those on their actual devices before relying on voice.

For avatar lifecycle changes, use local SDK doubles to assert that no avatar
starts before a successful session, and Stop/end/off/failure release the SDK,
peer connection, received tracks and timers. Exercise late startup, media failure
and the exact 30-second startup/120-second speech watchdog thresholds without
making billable Speech or Copilot requests. Verify text-only defaults, upfront
cost notices, and audio-only narration of terminal messages.
Assert one avatar startup/peer is reused across replies, verify each 15/34/45
second idle setting, and prove agent processing and playback suspend idle cleanup.
Idle expiry must retain the agent conversation and allow a later video reconnect.

## Privacy defaults

Azure Monitor automatic HTTP/dependency, console and browser tracing is
disabled; these can otherwise collect signed URLs or personal identifiers.
Only aggregate runtime/performance metrics are enabled by default. App Service
operational messages must not contain conversation data or credentials.
Do not turn on full request-body logging or browser session replay.

Power Automate secures connector and catalog-action inputs/outputs in run
history. App Service still sees the people-picker array to enforce policy,
then removes it before returning the catalog. SharePoint and Copilot Studio
have their own audit/transcript policies; the customer's privacy owner must
configure those independently. Log Analytics retention is 30 days in Bicep.
Document the operational owner and deletion/incident process before broad use.

## Scope of support

This is a Teams **personal tab**, not a separately registered Bot Framework
bot. The same HTTPS app can run standalone. The available live SDK transport
supports Standard-harness Copilot Studio agents; retaining other harness names
in SharePoint does not make those transports available.

Speech throttling is per verified user and in-memory. Multiple instances or
large rollouts require shared quota enforcement. A bearer token can be replayed
until expiry; do not claim sender-constrained or replay-proof authentication.
