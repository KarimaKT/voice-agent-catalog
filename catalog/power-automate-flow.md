# Maker-owned Power Automate catalog bridge

This flow lets the Teams app read the SharePoint agent catalog without asking
every user for Microsoft Graph `Sites.Read.All`.

## Security model

- The SharePoint **Get items** action uses the flow maker's connection.
- The HTTP trigger URL is a signed secret.
- Store that URL only in ignored `env/.env.dev.user` and Azure Key Vault.
- App Service calls the flow; the browser never receives the URL.
- The app validates its delegated bearer token, then sends the normalized email
  and object ID derived from verified claims.
- App Service repeats the Restricted-row policy check and removes
  `AllowedUsers` before responding to the browser.
- Copilot Studio sharing and runtime authorization remain the final security
  boundary. Do not put secrets in the catalog.

The Request trigger or organizational policy can require premium licensing.
Confirm licensing and data-loss-prevention policy with the customer.

## Create the flow

1. In Power Automate, create an **Instant cloud flow**.
2. Select **When an HTTP request is received**.
3. Use this request schema:

   ```json
   {
     "type": "object",
     "properties": {
       "operation": {
         "type": "string"
       },
       "userEmail": {
         "type": "string"
       },
       "userObjectId": {
         "type": "string"
       }
     }
   }
   ```

4. Add **SharePoint – Get items**.
5. Select the customer site and `Voice Agent Catalog` list.
6. Return catalog records only to the authenticated backend. The backend
   validates and filters `Enabled` and `Audience` and strips `AllowedUsers`
   before returning the public catalog. Do not expose the signed URL.
7. Determine catalog management UX using SharePoint site-user and list-effective
   permissions for the verified caller. Site administrators and users with
   `ManageLists` can manage the catalog. Set `configurationUrl` to the list/page URL
   only for those users. Return `false` and omit the URL for everyone else.
8. Add **Data Operation – Select**.
9. Set **From** to the filtered `value` output from **Get items**.
10. Map these output keys using the corresponding SharePoint dynamic content:

   | Output key | SharePoint column |
   | --- | --- |
   | `Id` | `ID` |
   | `Title` | `Title` |
   | `Description` | `Description` |
   | `EnvironmentId` | `EnvironmentId` |
   | `SchemaName` | `SchemaName` |
   | `Enabled` | `Enabled` |
   | `CompletionPhrase` | `CompletionPhrase` |
   | `EndsConversation` | `EndsConversation` |
   | `WelcomeMessage` | `WelcomeMessage` |
   | `Locale` | `Locale` |
   | `VoiceName` | `VoiceName` |
   | `AvatarCharacter` | `AvatarCharacter` |
   | `AvatarStyle` | `AvatarStyle` |
   | `Audience` | `Audience Value` from the Choice column |
   | `AllowedUsers` | `AllowedUsers` people-picker array (email/display name) |
   | `Harness` | `Harness` |

   Use the dynamic-content labels. This avoids relying on internal names such as
   `field_1` created by CSV import.

11. Add **Response**:
   - Status code: `200`
   - Header: `Content-Type` = `application/json`
   - Body:

   ```json
   {
     "agents": BODY_FROM_SELECT,
     "canManageCatalog": CAN_MANAGE_CATALOG,
     "configurationUrl": CONFIGURATION_URL_OR_NULL
   }
   ```

   Insert the **Select** body as dynamic content rather than typing
   `BODY_FROM_SELECT`.

12. Save the flow and copy its generated HTTP POST URL.

## Configure the app

Copy the secret template:

```powershell
Copy-Item env\.env.dev.user.example env\.env.dev.user
```

Set:

```dotenv
CATALOG_FLOW_URL=<signed-trigger-url>
```

Provision again. Bicep stores the URL in Key Vault as `catalog-flow-url` and
configures App Service with a Key Vault reference.

Before the flow is created, keep `CATALOG_FLOW_URL=disabled`. Production
connection remains unavailable; it does not bypass catalog policy with a default.

## Repeatable provisioning

[The template](catalog-flow.template.json) and
[PowerShell script](../scripts/provision-catalog-flow.ps1) reproduce the bridge.
Supply environment/site/list/connection identifiers and an ignored JSON field
map associating display names with actual internal names. The script saves
the flow ID and signed callback only inside a Git-ignored state file and updates
the same flow on reruns. It does not create or authorize connections.

`Audience` is a SharePoint Choice value (`Value`), not the full connector object.
Secure run-history inputs/outputs remain enabled. A failed permissions lookup
leaves `canManageCatalog=false`; it never grants the link by guessing.

## Test

POST this body from a trusted tool:

```json
{
  "operation": "listEnabledAgents",
  "userEmail": "user@example.com"
}
```

Confirm the response is:

```json
{
  "agents": [
    {
      "Id": "1",
      "Title": "Pat",
      "Enabled": true
    }
  ],
  "canManageCatalog": false
}
```

The real row must also contain all required identity, voice, and avatar fields.
Test an Everyone user, an allowed Restricted user, a disallowed user, a catalog
administrator, and a site administrator. Confirm only administrators receive a
safe `http`/`https` configuration URL.
