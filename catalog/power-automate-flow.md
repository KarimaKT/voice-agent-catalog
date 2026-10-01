# Maker-owned Power Automate catalog bridge

This flow lets the Teams app read the SharePoint agent catalog without asking
every user for Microsoft Graph `Sites.Read.All`.

## Security model

- The SharePoint **Get items** action uses the flow maker's connection.
- The HTTP trigger URL is a signed secret.
- Store that URL only in ignored `env/.env.dev.user` and Azure Key Vault.
- App Service calls the flow; the browser never receives the URL.
- Everyone who can use the app can see the normalized enabled-agent metadata
  returned by the flow. Do not put secrets in the catalog.

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
       }
     }
   }
   ```

4. Add **SharePoint – Get items**.
5. Select the customer site and `Pat Agent Catalog` list.
6. Add **Data Operation – Select**.
7. Set **From** to the `value` output from **Get items**.
8. Map these output keys using the corresponding SharePoint dynamic content:

   | Output key | SharePoint column |
   | --- | --- |
   | `Id` | `ID` |
   | `Title` | `Title` |
   | `Description` | `Description` |
   | `EnvironmentId` | `EnvironmentId` |
   | `SchemaName` | `SchemaName` |
   | `Enabled` | `Enabled` |
   | `CompletionPhrase` | `CompletionPhrase` |
   | `WelcomeMessage` | `WelcomeMessage` |
   | `Locale` | `Locale` |
   | `VoiceName` | `VoiceName` |
   | `AvatarCharacter` | `AvatarCharacter` |
   | `AvatarStyle` | `AvatarStyle` |

   Use the dynamic-content labels. This avoids relying on internal names such as
   `field_1` created by CSV import.

9. Add **Response**:
   - Status code: `200`
   - Header: `Content-Type` = `application/json`
   - Body:

   ```json
   {
     "agents": BODY_FROM_SELECT
   }
   ```

   Insert the **Select** body as dynamic content rather than typing
   `BODY_FROM_SELECT`.

10. Save the flow and copy its generated HTTP POST URL.

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

Before the flow is created, keep `CATALOG_FLOW_URL=disabled`. The default agent
continues to work while the shared catalog endpoint remains off.

## Test

POST this body from a trusted tool:

```json
{
  "operation": "listEnabledAgents"
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
  ]
}
```

The real row must also contain all required identity, voice, and avatar fields.
