# Architecture and security

## Design goal

The Teams application is a reusable voice shell. Conversation behavior belongs
to Copilot Studio agents, and agent presentation belongs to the SharePoint
catalog. This separation lets business owners publish or update agents without
rebuilding the Teams client.

## Runtime sequence

1. Teams loads the personal tab from Azure App Service.
2. The client initializes Teams JS and reads the Teams user's account context.
3. MSAL nested app authentication obtains the Teams user's token silently, or
   shows a one-time Copilot Studio consent dialog when required.
4. The client asks App Service for the catalog with a best-effort email hint.
5. App Service calls the signed Power Automate HTTP trigger from Key Vault.
6. The flow reads SharePoint using its maker-owned connector connection.
7. The user selects an agent.
8. The client obtains a delegated Power Platform token and invokes that agent.
9. The browser requests a short-lived Speech token from the same-origin broker.
10. App Service resolves the Speech key from Key Vault with managed identity.
11. Speech runs STT/TTS and, when enabled, establishes avatar WebRTC media.

STT uses the user's selected spoken input language. Catalog `Locale` remains
the output contract for the agent, TTS, and avatar. The input selection
defaults to that output locale.

## Why an Entra app registration is required

The registration identifies the application, not the person. It declares:

- the customer tenant;
- `brk-multihub://<deployed-domain>` for Teams nested app authentication;
- `/auth/callback` for browser diagnostics;
- Power Platform `CopilotStudio.Copilots.Invoke`.

Every user still authenticates as themselves. Their effective access is the
intersection of tenant consent and Copilot Studio sharing.

## Authentication UX

The catalog does not require a user Graph token. Power Automate uses the flow
maker's SharePoint connection and App Service keeps the signed trigger URL in
Key Vault. If the flow is unavailable, the default agent remains usable.

Teams NAA uses the identity already active in Teams, avoiding a separate account
login. If tenant policy permits user consent, the first connection can request
the Copilot Studio permission once. Later token acquisition is silent. The
browser-only callback page does not render React or an agent.

## Trust boundaries

- Browser tokens remain in MSAL session storage.
- Speech subscription key remains in Key Vault.
- Power Automate signed trigger URL remains in Key Vault.
- App Service receives the Speech key as a Key Vault reference.
- Browser Speech credentials are short-lived.
- The catalog connection is owned by Power Automate, not the browser.
- The Outlook connection is owned by Copilot Studio/Power Platform, not this
  web app.
- The app does not intentionally store messages or generated email HTML.
- Agent-provided URLs are accepted only for `http`/`https`; the React renderer
  uses elements and text nodes rather than unsafe HTML.

The catalog flow may filter `Restricted` rows using direct `AllowedUsers`
people-picker entries and may return an administrator-only configuration link.
The email hint is not proof of identity. Copilot Studio sharing and invocation
authorization are always the final security boundary.

## Harness compatibility

Catalog metadata accepts optional arbitrary `Harness` values for forward
compatibility. The current live transport uses the Copilot Studio Agents SDK
and has been verified only with the standard harness. GitHub Copilot/Copilot
Chat harnesses do not currently support Direct Line/Agents SDK transport, so
they are documented rather than rejected and require a future transport
adapter.

## POC limitations

The Speech token route has basic IP-based in-memory rate limiting but no bearer
token validation. Treat the deployment as a controlled POC. Add backend token
validation and distributed throttling before broader production use.

When `DEMO_MODE=true`, Pat and Morgan run deterministic client-side tools so the
demo remains available without Copilot consent. Morgan's session data mirrors
the SharePoint **Voice Agent Orders** list, but session changes are not
persisted. Pat can open a self-addressed Outlook draft but cannot silently send
it because the sample deliberately stores no user or Outlook credential.
