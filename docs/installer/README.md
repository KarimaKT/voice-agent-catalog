# Installer guide

This guide is for tenant administrators, platform owners, consultants, and
deployment engineers installing the complete Voice Agent Catalog.

## Table of contents

- [Choose a delivery channel](#choose-a-delivery-channel)
- [Choose an installation method](#choose-an-installation-method)
- [Prerequisites](#prerequisites)
- [Automated installation with the skill](#automated-installation-with-the-skill)
- [Manual installation](#manual-installation)
- [Import the sample agent](#import-the-sample-agent)
- [Deploy the web application](#deploy-the-web-application)
- [Use the standalone web application](#use-the-standalone-web-application)
- [Host the Node application elsewhere](#host-the-node-application-elsewhere)
- [Install or publish in Teams](#install-or-publish-in-teams)
- [Verify the installation](#verify-the-installation)
- [Handoff and operations](#handoff-and-operations)

## Choose a delivery channel

The same hosted application supports multiple channels:

| Channel | Intended use | What is published |
| --- | --- | --- |
| Standalone web application | Browser access, demos, and organizations that do not want a Teams package | Azure App Service URL |
| Self-hosted standalone web application | Customers with an existing Node.js hosting platform | Customer HTTPS URL |
| Teams personal app | Individual or pilot installation | Generated Teams app ZIP |
| Teams organizational catalog | Controlled tenant-wide distribution | Teams app ZIP submitted to Teams Admin Center |
| Public Teams store | Commercial/public distribution | Partner Center submission after Microsoft store validation |

The repository automates the first three. Public store publication requires
publisher verification, commercial listing information, support/privacy
material, and Microsoft certification outside this repository.

## Choose an installation method

### Recommended: run the included Agent Skill

Use a GitHub Copilot client, GitHub Copilot CLI, Scout, or another client that
supports repository Agent Skills. Open the repository and ask:

```text
Deploy the Voice Agent Catalog sample for my tenant.
```

The skill is stored at
`.github/skills/deploy-voice-agent-catalog/SKILL.md`. It collects one customer
value at a time, checks permissions, imports or creates the sample agent,
validates Azure, deploys the web application, packages Teams, and verifies the
result.

### Manual

Follow this guide and the command-oriented
[detailed installation procedure](../installation.md).

## Prerequisites

### Workstations

- Windows, macOS, or Linux with Git
- Node.js 22 or later and npm 10 or later
- Azure CLI
- Microsoft 365 Agents Toolkit CLI or VS Code extension
- Power Platform CLI when rebuilding the sample solution ZIP

### Customer permissions

- Azure subscription contributor or equivalent scoped permissions
- permission to create an Entra application registration
- Copilot Studio maker access in the target Power Platform environment
- permission to import Power Platform solutions
- SharePoint list creation rights
- Power Automate access with an HTTP Request trigger and SharePoint connector
- Teams custom app upload, or a Teams administrator for organizational rollout

### Accounts

Decide which customer account owns:

- the Copilot Studio agent;
- the Outlook connector action;
- the SharePoint catalog flow;
- the Azure deployment; and
- the Teams application.

Never put credentials, customer email addresses, tenant IDs, or signed flow URLs
in tracked files.

## Automated installation with the skill

The skill performs these phases:

1. customer and deployment intake;
2. prerequisites and account verification;
3. sample-agent import or agent creation;
4. Outlook action and recipient configuration;
5. SharePoint catalog and Power Automate bridge creation;
6. local ignored environment configuration;
7. Azure validation and what-if;
8. Azure, Entra, and Teams provisioning;
9. application deployment;
10. channel installation and end-to-end verification.

Interactive sign-ins remain customer-controlled. The skill must identify the
account to select before opening each sign-in and never asks for passwords,
tokens, MFA codes, or connector credentials.

## Manual installation

Clone the repository and create ignored local configuration:

```powershell
git clone <repository-url>
Set-Location voice-agent-catalog

Copy-Item env\.env.dev.example env\.env.dev
Copy-Item env\.env.dev.user.example env\.env.dev.user
```

Set the customer subscription, resource group, region, unique resource suffix,
Copilot environment/schema, signed catalog-flow URL, and demo-mode choice.

Install and build:

```powershell
npm install
npm run build
az bicep build --file infra\azure.bicep
```

Before deployment, validate App Service quota, Speech/avatar availability,
tenant policy, Bicep, ARM, what-if, and managed-identity Key Vault access.

## Import the sample agent

Import:

```text
copilot-studio/packages/pat-manager-handoff-solution.zip
```

[Download the sample solution ZIP](../../copilot-studio/packages/pat-manager-handoff-solution.zip).

In Power Apps or Copilot Studio:

1. Select the target environment.
2. Open **Solutions** and select **Import solution**.
3. Upload the ZIP.
4. Open **Pat Manager Handoff Sample**.
5. Add **Office 365 Outlook - Send an email (V2)**.
6. Configure the approved manager or shared-mailbox recipient by following
   `copilot-studio/outlook-action.md`.
7. Test insufficient and complete answers.
8. Confirm exactly one HTML email is sent.
9. Publish and share the agent.
10. Record its environment ID and schema name only in the ignored environment
    file.

The ZIP intentionally contains no Outlook connection, recipient, tenant ID, or
credentials. Connector authorization must occur in the destination environment.

## Deploy the web application

Authenticate Azure CLI and Agents Toolkit with the intended customer account:

```powershell
az login --tenant <tenant-id>
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login azure
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login m365
```

Provision and deploy:

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk provision `
  --env dev --folder . --interactive false

npx -y --package @microsoft/m365agentstoolkit-cli atk deploy `
  --env dev --folder . --interactive false
```

Provisioning creates or updates App Service, Speech, Key Vault, Application
Insights, Log Analytics, Entra, and the Teams app registration.

Provisioning first persists the app API scope using `aad.api.manifest.json`,
then applies preauthorization using `aad.manifest.json`. Preserve this order.
The deployment lifecycle runs tests and typecheck before build/zip deployment.

For simple maker onboarding, install the
[URL-registration flow](../../catalog/agent-registration-flow.template.json)
with an authenticated maker-owned Dataverse connection. On the new-item form,
show only the agent URL; keep profile/access fields on the edit form and default
new entries to disabled. The flow resolves metadata only in the configured
environment. Configure additional environments explicitly rather than inventing
schema names or routing arbitrary URLs.

The product and Teams UX name is **Voice Agent Catalog**. For an existing
installation, keep the generated `TEAMS_APP_ID`, Entra client/object IDs, and
resource identifiers in the environment files and rerun provisioning so the
manifest updates in place. Do not delete and recreate registrations merely to
apply the name.

## Use the standalone web application

Open the deployed endpoint in a supported browser:

```text
https://<app-name>.azurewebsites.net/
```

The root redirects to the voice-agent experience. Outside Teams, the app uses
browser MSAL instead of Teams nested app authentication. The Entra registration
must contain the deployed `/auth/callback` SPA redirect.

The standalone site must be served over HTTPS for microphone and WebRTC avatar
features. Browser users still need access to every selected Copilot Studio
agent.

## Host the Node application elsewhere

The runtime is a standard Node.js 22 application and is not coupled to App
Service. A customer can deploy the built repository to another HTTPS-capable
Node host:

```powershell
npm ci
npm run build
npm start
```

The host must provide these runtime environment values securely:

- `TENANT_ID`
- `AAD_APP_CLIENT_ID`
- `COPILOT_ENVIRONMENT_ID`
- `COPILOT_SCHEMA_NAME`
- `SPEECH_REGION`
- `SPEECH_ENDPOINT`
- `SPEECH_KEY`
- the required production catalog flow URL, locale, voice, avatar, and optional demo settings

Use the destination platform's secret manager for `SPEECH_KEY` and
`CATALOG_FLOW_URL`. Do not place them in an image, repository, client bundle, or
plain-text deployment manifest.

Register the final origin and `/auth/callback` in Entra. If the same host also
serves the Teams package, update `TAB_ENDPOINT`, `TAB_DOMAIN`, and the NAA
`brk-multihub://<domain>` redirect before rebuilding the package.

## Install or publish in Teams

### Personal or pilot installation

Install the generated package:

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk install `
  --file-path appPackage\build\appPackage.dev.zip `
  --scope Personal --interactive false
```

### Organizational catalog

Use the repository publish stage:

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk publish `
  --env dev --folder . --interactive false
```

This submits the generated package for review in Teams Admin Center. A Teams
administrator must approve it and apply the appropriate app permission, setup,
or assignment policies.

Do not commit generated tenant-specific Teams packages.

## Verify the installation

Verify:

- `/api/health` returns `{"status":"ok"}`;
- the root URL and `/tabs/home` load;
- security headers are present;
- Key Vault references resolve;
- the catalog flow returns enabled agents;
- Teams reuses the current Teams identity;
- browser sign-in returns through `/auth/callback`;
- typed and spoken turns reach the selected agent;
- the selected voice and avatar are applied;
- Pat sends exactly one HTML email after successful completion; and
- Application Insights contains operational failures but no transcripts.

See [operations and troubleshooting](../operations.md) for detailed diagnostics.
Complete [all release gates](../release-validation.md) before organizational
catalog submission. A successful maker test-panel interview is not a substitute
for published-agent verification through the deployed app.

## Handoff and operations

Provide the customer:

- standalone application URL;
- Teams package or organizational catalog status;
- Copilot Studio environment and agent name;
- SharePoint catalog URL;
- Azure resource group and region;
- access/consent actions still required;
- cost ownership; and
- cleanup process.

Do not include secrets, signed URLs, connection identifiers, or recipient
addresses in the handoff document.
