# Customer installation

## 1. Choose the customer owner

Use one work account that can:

- create or manage Copilot Studio agents;
- authorize the Outlook connection;
- create a SharePoint list;
- create Entra and Teams applications; and
- deploy to the selected Azure subscription.

Use that same account for all installation sign-ins unless the customer
explicitly separates duties.

## 2. Prepare the Copilot Studio agent

Create a standard Microsoft-authenticated agent or import the tenant-neutral
solution when supplied. Pat's reference instructions are under
`copilot-studio/`.

Publish the agent and record:

- environment GUID;
- schema name;
- display name and description;
- welcome and completion phrases; and
- intended users/security group.

For Pat, configure Office 365 Outlook `SendEmailV2` with the customer's manager
or shared-mailbox recipient. Destination credentials are never included in the
repository.

## 3. Create the catalog

Create a Microsoft List from `catalog/Pat Agent Catalog.csv`. Replace placeholder
agent identifiers and apply the selected voice/avatar values. Grant users read
access and maintainers edit access.

## 4. Create the connector bridge

Follow `catalog/power-automate-flow.md`. Save the flow with the deployment
owner's SharePoint connection and copy the signed HTTP trigger URL.

## 5. Configure local deployment values

Copy:

```powershell
Copy-Item env\.env.dev.example env\.env.dev
Copy-Item env\.env.dev.user.example env\.env.dev.user
```

Set subscription, resource group, region, suffix, Copilot environment/schema,
and the signed flow URL. Both customer files are ignored by Git.

## 6. Validate prerequisites

```powershell
node --version
npm --version
az version
npx -y --package @microsoft/m365agentstoolkit-cli atk --version
npm install
npm run build
az bicep build --file infra\azure.bicep
```

Confirm:

- App Service B1 quota;
- Speech S0/avatar support in the region;
- Azure write permission;
- Entra app creation permission;
- custom Teams app policy; and
- Power Platform capacity.

## 7. Authenticate

Before opening each login, tell the installer which account to select.

```powershell
az login --tenant <tenant-guid>
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login azure
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login m365
```

Azure CLI and the two Agents Toolkit services have separate secure token caches.

## 8. Provision and deploy

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk provision `
  --env dev --folder . --interactive false

npx -y --package @microsoft/m365agentstoolkit-cli atk deploy `
  --env dev --folder . --interactive false
```

If provisioning fails after app IDs were created, keep the generated values,
fix the root cause, validate, and rerun. Provisioning is idempotent.

## 9. Consent and sharing

The Entra app requests Power Platform `CopilotStudio.Copilots.Invoke`. Users can
consent once when tenant policy permits; otherwise an administrator must grant
consent. In Teams, MSAL nested app authentication reuses the current Teams work
identity, so users do not complete a separate account login.

Share every enabled agent with the intended Teams users. The flow maker must
retain read access to the SharePoint list.

## 10. Install Teams package

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk install `
  --file-path appPackage\build\appPackage.dev.zip `
  --scope Personal --interactive false
```

For wider controlled rollout, upload the ZIP in Teams Admin Center and use app
permission/setup policies.

## 11. Verify

Follow the verification list in the main README and `docs/operations.md`.
