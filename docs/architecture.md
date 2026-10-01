# Architecture and security

## Design goal

The Teams application is a reusable voice shell. Conversation behavior belongs
to Copilot Studio agents, and agent presentation belongs to the SharePoint
catalog. This separation lets business owners publish or update agents without
rebuilding the Teams client.

## Runtime sequence

1. Teams loads the personal tab from Azure App Service.
2. The client initializes Teams JS and reads the Teams user's login hint.
3. MSAL signs in the existing work user through a dedicated blank callback.
4. The client obtains a delegated Graph token and reads catalog columns/items.
5. The user selects an agent.
6. The client obtains a delegated Power Platform token and invokes that agent.
7. The browser requests a short-lived Speech token from the same-origin broker.
8. App Service resolves the Speech key from Key Vault with managed identity.
9. Speech runs STT/TTS and, when enabled, establishes avatar WebRTC media.

## Why an Entra app registration is required

The registration identifies the application, not the person. It declares:

- the customer tenant;
- `/auth/callback` as the SPA callback;
- Microsoft Graph `Sites.Read.All`; and
- Power Platform `CopilotStudio.Copilots.Invoke`.

Every user still authenticates as themselves. Their effective access is the
intersection of tenant consent, SharePoint access, and Copilot Studio sharing.

## Authentication UX

Graph and Power Platform issue different audience tokens. The app cannot request
both resources in one OAuth token. It signs in with the Graph scope first so it
can load the catalog, then requests the selected agent's Power Platform scope.
With admin pre-consent, later token acquisition is silent. Without pre-consent,
first use can show separate consent interactions.

The callback page does not render React or an agent. This prevents nested copies
of the app inside authentication windows.

## Trust boundaries

- Browser tokens remain in MSAL session storage.
- Speech subscription key remains in Key Vault.
- App Service receives the Speech key as a Key Vault reference.
- Browser Speech credentials are short-lived.
- The Outlook connection is owned by Copilot Studio/Power Platform, not this
  web app.
- The app does not intentionally store messages or generated email HTML.

## POC limitations

The Speech token route has basic IP-based in-memory rate limiting but no bearer
token validation. Treat the deployment as a controlled POC. Add backend token
validation and distributed throttling before broader production use.
