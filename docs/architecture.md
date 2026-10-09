# Architecture and security

## Design goal

The Teams application is a reusable voice shell. Conversation behavior belongs
to Copilot Studio agents, and agent presentation belongs to the SharePoint
catalog. This separation lets business owners publish or update agents without
rebuilding the Teams client.

## Runtime sequence

1. Teams loads the personal tab from Azure App Service.
2. The client initializes Teams JS and reads the Teams user's account context.
3. MSAL nested app authentication obtains a delegated app API token silently,
   or shows one consent prompt when required.
4. The client sends that bearer token to App Service for the catalog.
5. App Service calls the signed Power Automate HTTP trigger from Key Vault.
6. The flow reads SharePoint using its maker-owned connector connection.
7. The user selects an agent.
8. The client separately obtains a delegated Power Platform token and invokes
   that agent.
9. The browser requests a short-lived Speech token using the app API bearer.
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
- the app-owned `access_as_user` delegated API scope, preauthorized for the
  application and Microsoft Teams web/desktop clients.

Every user still authenticates as themselves. Their effective access is the
intersection of tenant consent and Copilot Studio sharing.

## Authentication UX

The catalog does not require a Graph token. Power Automate uses the flow
maker's SharePoint connection and App Service keeps the signed trigger URL in
Key Vault. Production connection fails closed if the flow is unavailable or
unconfigured. Only an explicitly labelled demo uses local simulated agents.

Teams NAA uses the identity already active in Teams, avoiding a separate account
login. Tenant admin consent can cover `access_as_user` and the separate Copilot
Studio permission. Later token acquisition is silent. The
browser-only callback page does not render React or an agent.
Standalone interaction uses same-window MSAL redirect rather than relying on
browser popup windows. The main application processes the redirect result and
reloads the authenticated catalog without starting a conversation. Teams keeps
the nested broker path; no browser redirect is started inside Teams.

MSAL Browser v5 requires the dedicated `/auth/callback` page to load
`@azure/msal-browser/redirect-bridge` and broadcast the authorization response.
Vite builds this as a separate entry; the host sends the built HTML, not a plain
placeholder page. The callback must not use COOP headers. CSP permits the
same-origin callback and the Microsoft authority iframe without loosening the
script policy.

## Trust boundaries

- Browser tokens remain in MSAL session storage.
- App Service validates issuer, tenant, audience, expiry, signature, and
  `access_as_user` scope against the tenant's Entra signing keys.
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

The backend derives normalized email/UPN and object ID only from verified
claims. It validates enabled rows, applies direct `AllowedUsers` equality for
`Restricted` rows, and never returns `AllowedUsers` to the browser. Copilot
Studio sharing and invocation authorization remain the final boundary.

## Harness compatibility

Catalog metadata accepts arbitrary `Harness` values for forward compatibility.
Only `standard` is currently invokable. Other rows remain visible but disabled
with an unavailable-transport label; they require a future transport adapter.

## Release and lifecycle

`aad.api.manifest.json` registers the delegated scope before
`aad.manifest.json` refers to it in preauthorization. Entra rejects references
to scopes not yet persisted, so provisioning uses these two ordered updates.
Generated IDs are preserved on reruns.

`src/auth.ts` owns token verification, `src/catalog-policy.ts` owns server-side
row validation/visibility, and the React client consumes only sanitized rows.
The flow is a connector boundary, not an alternative source of caller identity.
The UI sends voice-interface context at the start of every real conversation;
this requests output-language/style adaptation but cannot override agent rules.
Silent startup uses the SDK's typed start response: its conversation ID may be
returned in response metadata even when no start activities are emitted. The
subsequent interface message and user turns use that ID and remain streamed.

See [release gates and privacy defaults](release-validation.md).

## POC limitations

The Speech route also has per-user in-memory rate limiting. Use
distributed throttling before broad multi-instance production rollout.

When `DEMO_MODE=true`, Pat and Morgan run deterministic client-side tools so the
demo can exercise conversation behavior without Copilot consent. Speech still
requires an authenticated app API token. Morgan's session data mirrors
the SharePoint **Voice Agent Orders** list, but session changes are not
persisted. Pat can open a self-addressed Outlook draft but cannot silently send
it because the sample deliberately stores no user or Outlook credential.
