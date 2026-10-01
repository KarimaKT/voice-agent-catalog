---
name: deploy-voice-agent-catalog
description: |
  Collect customer-specific configuration and deploy the Voice Agent Catalog
  sample end to end: Copilot Studio reference agent and actions, SharePoint
  agent catalog, Entra application, Azure resources, Teams personal app, and
  verification. Use when a customer asks to install, configure, provision,
  deploy, publish, or update this sample.
license: MIT
---

# Deploy Voice Agent Catalog

Use this skill only from the root of the Voice Agent Catalog repository.
Treat the deployment as an interactive business-user installation, not as a
generic infrastructure exercise.

## Outcome

Complete all of these surfaces:

1. A Microsoft-authenticated standard Copilot Studio agent.
2. One Office 365 Outlook `SendEmailV2` action that sends the final HTML handoff.
3. A published agent shared with the intended testers.
4. A SharePoint/Microsoft List catalog containing the agent's voice and avatar
   identity.
5. A maker-owned Power Automate bridge that reads the catalog without delegated
   Graph permission for app users.
6. A single-tenant Entra SPA and tenant-specific Teams personal app.
7. Azure App Service, Speech S0, Key Vault, Application Insights, and Log
   Analytics resources.
8. A deployed and verified application and an installable Teams package.

The Teams component is a personal tab that calls the Copilot Studio agent. Do
not describe or provision a separate Bot Framework bot unless the repository
architecture has intentionally changed.

## Safety and privacy rules

- Never ask for or store passwords, refresh tokens, connector credentials,
  Speech keys, client secrets, or MFA codes.
- Tell the user which account to select before opening every interactive login.
- Keep tenant IDs, subscription IDs, emails, resource names, and generated app
  IDs in ignored local files only.
- Never place customer-specific values in tracked examples, documentation,
  screenshots, or the generic catalog CSV.
- Use Teams nested app authentication for delegated Copilot Studio access.
- Use the maker-owned Power Automate SharePoint connection for catalog access;
  do not request delegated Graph SharePoint permissions from app users.
- Use the App Service managed identity and a Key Vault reference for the Speech
  key.
- Do not log interview answers, generated email HTML, access tokens, or Speech
  credentials.
- Require explicit confirmation before deleting or replacing resources.
- Stop on failed deployment, permission, consent, agent publish, email, or
  catalog validation. Do not return success-shaped fallbacks.

## Interaction rules

Ask one question at a time. Prefer choices and mark the recommended answer.
Persist completed answers in `.azure/customer-deployment.json`; `.azure/` is
ignored by Git. Do not ask again for a value that can be discovered safely from
the signed-in account or existing project files.

When an authentication window is required, first say:

> Sign in with `<customer account>`. Do not select another cached account.

After authentication, verify the returned account and tenant before continuing.
If the wrong account was selected, sign it out and retry immediately.

## Phase 1: Intake

Collect or discover these values:

### Customer and access

- Deployment owner email.
- Microsoft Entra tenant ID.
- Azure subscription name and ID.
- Azure region that supports Speech S0 real-time avatars and has App Service B1
  quota.
- Names or email addresses of up to three initial testers.

### Resource naming

- Dedicated resource group name.
- Globally unique lowercase resource suffix, 4-20 characters.
- Environment name, normally `dev` for a POC.

### Copilot Studio

- Target Copilot Studio environment URL or environment ID.
- New agent or existing compatible standard agent.
- Agent display name.
- Agent schema name after creation or import.
- Manager or shared-mailbox recipient.
- Maker-owned or user-owned Outlook connection policy.
- Exact completion phrase.
- Welcome message.

### SharePoint catalog

- SharePoint hostname.
- Server-relative site path.
- Catalog list display name.
- Catalog editors.
- Flow maker account and SharePoint connector connection.
- Power Automate Request trigger licensing and DLP approval.

### Speech identity

- Recognition/synthesis locale.
- Azure neural voice name.
- Avatar character.
- Avatar style supported by that character.

Recommend these POC defaults:

- Locale: `en-US`
- Voice: `en-US-AvaMultilingualNeural`
- Avatar: `lisa`
- Style: `casual-sitting`
- Completion phrase:
  `Thank you. Your manager handoff summary has been emailed.`

## Phase 2: Prerequisites and sign-in

Verify:

```powershell
node --version
npm --version
az version
npx -y --package @microsoft/m365agentstoolkit-cli atk --version
```

Use Node.js 22 or later. Use an isolated npm cache if the shared npm cache
reports lock or integrity errors.

Authenticate Azure CLI and Agents Toolkit separately because they have separate
token caches:

```powershell
az login --tenant <tenant-id>
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login azure
npx -y --package @microsoft/m365agentstoolkit-cli atk auth login m365
```

After each login, verify the account. Confirm the target subscription and
region with the user before provisioning.

## Phase 3: Copilot Studio agent

Prefer importing the tenant-neutral unmanaged solution when the repository
contains one under `copilot-studio/solution/`. If no solution package is
present, create the agent with the included tenant-neutral files:

- `copilot-studio/agent-instructions.txt`
- `copilot-studio/outlook-action.md`
- `copilot-studio/README.md`

When Copilot Studio authoring skills are available, use them in this order:

1. Create or clone the standard-harness agent.
2. Apply the agent instructions.
3. Add and configure Office 365 Outlook `SendEmailV2`.
4. Validate the agent definition.
5. Publish the agent.

Otherwise, guide or automate the equivalent Copilot Studio portal steps.

Required agent behavior:

1. Ask for name and role.
2. Ask for the most important responsibility.
3. Ask for one key process, document, or system.
4. Ask for one open item, risk, or recommendation.
5. Ask one concise follow-up when an answer is insufficient.
6. Call Outlook exactly once after all four topics are sufficient.
7. Send an Outlook-compatible HTML handoff to the configured recipient.
8. Wait for confirmed action success.
9. Say the exact completion phrase only after successful delivery.

Use Microsoft authentication. Configure the recipient during installation; do
not allow an interviewee to choose an arbitrary recipient unless the customer
explicitly requires and secures that behavior.

Before continuing, test:

- one insufficient answer and follow-up;
- all four topics;
- exactly one email;
- formatted HTML in Outlook;
- exact final completion phrase; and
- published-agent access for each tester.

Record the published environment ID and schema name locally.

## Phase 4: SharePoint agent catalog

Create the list from `catalog/Pat Agent Catalog.csv` on the selected SharePoint
site, or update an existing compatible list. The required display-name columns
are:

- `Title`
- `Description`
- `EnvironmentId`
- `SchemaName`
- `Enabled`
- `CompletionPhrase`
- `WelcomeMessage`
- `Locale`
- `VoiceName`
- `AvatarCharacter`
- `AvatarStyle`

Do not assume SharePoint internal field names equal display names. CSV import
can produce names such as `field_1`; map Power Automate **Select** keys with the
SharePoint dynamic-content labels.

Create or update one enabled row for the published agent. Apply the selected
locale, voice, avatar character, and avatar style. Grant the flow maker read
access and catalog maintainers edit access.

Read the row back and verify all eleven fields before continuing.

Create the maker-owned flow exactly as documented in
`catalog/power-automate-flow.md`. Save it, copy the signed trigger URL, and test
that it returns an `agents` array. Treat the URL as a secret.

## Phase 5: Local environment

Copy `env/.env.dev.example` to the selected ignored environment file and set:

```dotenv
TEAMSFX_ENV=<environment-name>
APP_NAME_SUFFIX=<environment-name>
AZURE_SUBSCRIPTION_ID=<subscription-guid>
AZURE_RESOURCE_GROUP_NAME=<resource-group>
AZURE_LOCATION=<validated-region>
RESOURCE_SUFFIX=<unique-suffix>
COPILOT_ENVIRONMENT_ID=<published-environment-guid>
COPILOT_SCHEMA_NAME=<published-schema-name>
```

Copy `env/.env.dev.user.example` to the matching ignored user file and set:

```dotenv
CATALOG_FLOW_URL=<signed-power-automate-trigger-url>
DEMO_MODE=false
```

Use `CATALOG_FLOW_URL=disabled` only for a staged deployment before the flow
exists. The configured default agent remains available, but `/api/catalog`
stays disabled until the signed URL is supplied and provisioning is rerun.

For a time-critical POC when delegated consent is blocked, set
`DEMO_MODE=true`. This runs the local four-question Pat demonstration and the
deterministic Morgan order tool with Speech/avatar but intentionally bypasses
Copilot Studio. In Teams, Pat uses the current user's login hint to open a
self-addressed Outlook draft; the user must select **Send**. The UI and final
summary must clearly identify those limitations. Set demo mode back to `false`
before verifying real agent invocation, SharePoint persistence, or automatic
Outlook delivery.

Leave generated Entra, Teams, endpoint, and resource ID values empty before the
first provision. Confirm the environment file is ignored by Git.

## Phase 6: Build and Azure validation

Run:

```powershell
npm install
npm run build
az bicep build --file infra/azure.bicep
```

Follow the repository's Azure preparation and validation workflow. Validation
must include:

- authenticated subscription;
- Bicep build and lint;
- ARM validation;
- what-if with no unexpected deletions;
- App Service quota in the chosen region;
- Speech avatar availability;
- tenant policy review; and
- static managed-identity/Key Vault access verification.

Do not deploy until `.azure/deployment-plan.md` has status `Validated`.

## Phase 7: Provision and deploy

Use the repository lifecycle:

```powershell
npx -y --package @microsoft/m365agentstoolkit-cli atk provision `
  --env <environment-name> --folder . --interactive false

npx -y --package @microsoft/m365agentstoolkit-cli atk deploy `
  --env <environment-name> --folder . --interactive false
```

Provisioning is idempotent. If a stage fails after creating Entra or Teams app
IDs, preserve the generated environment values, fix the root cause, validate
the change, and rerun. Do not create replacement registrations unnecessarily.

If tenant policy requires Key Vault purge protection, keep it enabled. If a
region has zero B1 quota, select another Speech-avatar-supported region and
revalidate instead of silently changing the architecture.

## Phase 8: Consent, package, and installation

Verify the Entra SPA has:

- Power Platform `CopilotStudio.Copilots.Invoke`;
- `brk-multihub://<deployed-domain>` as a SPA redirect for Teams NAA;
- the deployed `/auth/callback` SPA redirect URI; and
- single-tenant sign-in.

The Teams client must initialize Teams JS before MSAL and use nested app
authentication. Verify silent token acquisition first. If user consent is
allowed, first use can show a one-time Copilot Studio permission dialog without
a separate account login. Grant tenant admin consent only when customer policy
blocks user consent. Never claim consent succeeded without verifying the token
or service principal.

Build and validate the Teams package with Agents Toolkit. Install the generated
tenant-specific ZIP for the deployment owner, then share or assign it to the
initial testers according to tenant policy.

## Phase 9: End-to-end verification

Verify all of these:

1. `GET /api/health` returns `{"status":"ok"}`.
2. `/tabs/home` returns HTTP 200 and the Content Security Policy header.
3. `/api/config` contains the expected default agent and catalog-enabled state.
   If demo mode is selected, also verify `demoMode: true`.
4. Key Vault references resolve and neither secret reaches the browser.
5. `/api/catalog` returns the maker-owned flow's normalized `agents` array.
   If `CATALOG_FLOW_URL=disabled`, verify it returns 404 and the default agent
   remains available instead.
6. Teams uses the current work identity without a separate login prompt.
7. First-use consent, if required, returns to the app without a nested app
   window; subsequent token acquisition is silent.
8. Agent dropdown and catalog identity are correct.
9. Typed conversation reaches the published agent, or completes the clearly
   labeled local four-question flow when `DEMO_MODE=true`.
10. STT recognizes microphone input using the selected locale.
    Verify managed voice mode starts with one action, submits after end-of-turn
    silence, pauses while the agent speaks, reopens automatically afterward,
    and stops when **Stop conversation** is selected.
11. Audio TTS uses the selected neural voice.
12. Avatar mode uses the selected character/style, with audio-only fallback
    when relay credentials are unavailable.
13. Completing the real-agent interview sends exactly one HTML email. In demo
    mode, verify that the UI explicitly says email was not sent.
14. The UI displays the selected agent's configured completion phrase.
15. Application Insights contains health and failure telemetry but no transcript
    or generated HTML content.
16. Managed-identity Key Vault access is present in live Azure state.

If any check fails, fix the root cause and repeat the smallest relevant
provision, deploy, or verification step.

## Phase 10: Handoff

Provide the customer:

- fully qualified `https://` application endpoint;
- Teams package path;
- Copilot Studio agent name and environment;
- SharePoint list URL;
- deployed Azure resource group and region;
- tester access status;
- cost and cleanup guidance; and
- any tenant-admin action still required.

Do not include tokens, keys, private connection identifiers, or the manager
email in the handoff report.

Before publishing the repository, scan tracked content for customer tenant IDs,
subscription IDs, environment IDs, agent IDs, emails, absolute user paths, and
generated packages.
