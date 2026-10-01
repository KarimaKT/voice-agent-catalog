# Microsoft List agent catalog

The app can populate its agent dropdown from a Microsoft List stored on a
SharePoint site. Pat remains the fallback when no SharePoint catalog is
configured.

## Fastest setup for a business user

1. Open the SharePoint site that should own the catalog.
2. Select **New → List → From CSV**.
3. Upload `Pat Agent Catalog.csv`.
4. Name the list `Pat Agent Catalog`.
5. Confirm these columns are created:

   | Column | Recommended type |
   | --- | --- |
   | Title | Single line of text |
   | Description | Multiple lines of text |
   | EnvironmentId | Single line of text |
   | SchemaName | Single line of text |
   | Enabled | Yes/No, or single line of text containing `TRUE`/`FALSE` |
   | CompletionPhrase | Multiple lines of text |
   | WelcomeMessage | Multiple lines of text |
   | Locale | Single line of text |
   | VoiceName | Single line of text |
   | AvatarCharacter | Single line of text |
   | AvatarStyle | Single line of text |

6. Replace the placeholder environment ID and schema name in the Pat row with
   values from the published Copilot Studio agent.
7. Give catalog editors **Edit** access to the list.
8. Give app users **Read** access to the list and permission to use every agent
   you enable.
9. Create the maker-owned connector bridge in
   [`power-automate-flow.md`](power-automate-flow.md).
10. Store its signed trigger URL in ignored `env/.env.dev.user`.
11. Provision and deploy the app.

## Add an agent

Create a new list row:

- **Title:** dropdown label
- **Description:** short UI subtitle
- **EnvironmentId:** Copilot Studio environment GUID
- **SchemaName:** published agent schema name
- **Enabled:** true
- **CompletionPhrase:** exact deterministic final success sentence
- **WelcomeMessage:** first prompt shown when the agent emits no start message
- **Locale:** Speech recognition and synthesis locale, such as `en-US`
- **VoiceName:** Azure neural voice, such as `en-US-AvaMultilingualNeural`
- **AvatarCharacter:** Azure avatar character, such as `lisa`
- **AvatarStyle:** a style supported by that character, such as
  `casual-sitting`

The app reads the list each time a user connects. Adding, disabling, or editing
an agent does not require rebuilding or reinstalling the Teams app.

## Order-management demo list

`Voice Agent Orders.csv` defines the sample data used by Morgan. Import it into
the same SharePoint site as a list named **Voice Agent Orders**. The no-login
demo uses a deterministic in-session mirror of these rows, so it does not
depend on delegated SharePoint consent. For production, replace that demo tool
with a maker-owned Power Automate SharePoint connector action.

## Connector and access model

The app does not request Microsoft Graph SharePoint permissions. App Service
calls a signed Power Automate HTTP trigger, and the flow reads the list with the
maker-owned SharePoint connector connection.

The signed trigger URL is stored in Key Vault and never returned to the browser.
Do not put credentials or secrets in catalog rows. Every authorized app user can
receive the enabled-agent metadata returned by the flow.

## Validation

Enabled rows must have:

- a valid environment GUID;
- a schema name containing only letters, numbers, `_`, `.`, or `-`;
- a nonempty title; and
- a locale, voice name, avatar character, and avatar style supported by the
  deployed Azure Speech region; and
- access granted to the signed-in user.

Invalid rows produce a visible configuration error rather than silently routing
the user to an unexpected agent.
