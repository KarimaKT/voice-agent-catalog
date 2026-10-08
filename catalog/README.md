# Microsoft List agent catalog

The app can populate its agent dropdown from a Microsoft List stored on a
SharePoint site. Pat remains the fallback when no SharePoint catalog is
configured.

## Fastest setup for a business user

1. Open the SharePoint site that should own the catalog.
2. Select **New → List → From CSV**.
3. Upload `Voice Agent Catalog.csv`.
4. Name the list `Voice Agent Catalog`.
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
   | Audience | Choice: `Everyone`, `Restricted` |
   | AllowedUsers | Person or Group, people only, multiple selections |
   | Harness | Single line of text (optional) |

6. Replace the placeholder environment ID and schema name in the Pat row with
   values from the published Copilot Studio agent.
7. Give catalog editors **Edit** access to the list.
8. Share each Copilot Studio agent with its intended runtime users. The flow's
   audience filtering improves dropdown UX; Copilot Studio sharing remains the
   final security boundary.
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
- **Locale:** agent output and Speech synthesis locale, such as `en-US`
- **VoiceName:** Azure neural voice, such as `en-US-AvaMultilingualNeural`
- **AvatarCharacter:** Azure avatar character, such as `lisa`
- **AvatarStyle:** a style supported by that character, such as
  `casual-sitting`
- **Audience:** `Everyone` or `Restricted`
- **AllowedUsers:** direct users allowed to see a `Restricted` row
- **Harness:** optional descriptive value. Use `standard` for the currently
  verified transport; unknown future values are retained rather than rejected.

Users separately choose their spoken input language in the app. It defaults to
the selected agent's output `Locale`; STT uses the input choice while agent
responses, TTS, and avatar speech use `Locale`.

After CSV import, replace the text `AllowedUsers` column with a SharePoint
**Person or Group** column configured for **People only** and multiple
selections. The flow must emit each selected person's email address. Only
direct selections are evaluated; this template does not expand groups.

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
Do not put credentials or secrets in catalog rows.

The backend sends a best-effort user email to the flow for UX filtering. It is
not an authorization credential and must not be treated as one. The flow
returns `canManageCatalog` and `configurationUrl`; the app displays the
configuration link only when the flow says the current user is an administrator
or site administrator. Do not publish a raw list URL through app environment
variables.

**Copilot Studio agent sharing and runtime authorization are the final security
boundary.** Catalog filtering does not grant access and must not replace
sharing, authentication, or connector authorization.

## Validation

Enabled rows must have:

- a valid environment GUID;
- a schema name containing only letters, numbers, `_`, `.`, or `-`;
- a nonempty title; and
- a locale, voice name, avatar character, and avatar style supported by the
  deployed Azure Speech region; and
- access granted to the signed-in user.

`Audience=Restricted` rows also require at least one valid `AllowedUsers`
email. The app hides restricted rows when it cannot match the best-effort user
email, but the flow should perform the same filtering before returning data.

Invalid rows produce a visible configuration error rather than silently routing
the user to an unexpected agent.
