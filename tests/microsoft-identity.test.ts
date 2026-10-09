import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import { runInNewContext } from "node:vm";

const tenantId = "11111111-1111-1111-1111-111111111111";
const guest = {
  tenantId,
  homeAccountId: "external-home-object.external-home-tenant",
  username: "guest@example.com",
};
class InteractionRequiredAuthError extends Error {}

// Execute the actual identity implementation with an MSAL test double, without
// installing dependencies or accessing the frozen deployment's node_modules.
async function identityHarness(teams = true) {
  const filters: unknown[] = [];
  const activated: unknown[] = [];
  let popups = 0;
  let silentCalls = 0;
  let nestedFactories = 0;
  let standardFactories = 0;
  const msal = {
    getActiveAccount: () => guest,
    getAccount: (filter: unknown) => { filters.push(filter); return guest; },
    setActiveAccount: (account: unknown) => activated.push(account),
    acquireTokenSilent: async () => {
      silentCalls++;
      return { accessToken: "silent-token", account: guest };
    },
    acquireTokenPopup: async () => {
      popups++;
      return { accessToken: "popup-token", account: guest };
    },
  };
  const exportsByModule: Record<string, Record<string, unknown>> = {
    "@azure/msal-browser": {
      InteractionRequiredAuthError,
      createNestablePublicClientApplication: async () => { nestedFactories++; return msal; },
      createStandardPublicClientApplication: async () => { standardFactories++; return msal; },
    },
    "@microsoft/agents-activity": { Activity: class {}, ActivityTypes: {} },
    "@microsoft/agents-copilotstudio-client": {
      ConnectionSettings: class {}, CopilotStudioClient: class {}, ScopeHelper: {},
    },
    "./agent-behavior": { getHarnessSupport: () => ({ supported: true }), voiceInterfaceInstructions: () => "" },
    "./adaptive-card-policy": { safeHttpUrl: () => undefined },
  };
  const source = readFileSync(new URL("../src/Tab/copilot.ts", import.meta.url), "utf8");
  const script = stripTypeScriptTypes(source, { mode: "transform" })
    .replace(/^import[\s\S]*?;\r?$/gm, "")
    .replace(/^export /gm, "");
  const Identity = runInNewContext(`${script}\nMicrosoftIdentity;`, {
    window: { location: { origin: "https://app.example.com" } },
    ...Object.assign({}, ...Object.values(exportsByModule)),
  }) as new (
    config: Record<string, unknown>, context: Record<string, unknown>,
  ) => { acquireApiToken(): Promise<string> };
  const teamsContext = {
    isTeamsHosted: teams, supportsNestedAuth: teams,
    loginHint: "guest@example.com", tenantId,
  };
  const config = { clientId: "client", tenantId, apiScope: "api://client/access_as_user" };
  return {
    identity: new Identity(config, teamsContext),
    createIdentity: (overrides: Record<string, unknown>) => new Identity(config, { ...teamsContext, ...overrides }),
    msal, filters, activated,
    counts: () => ({ popups, silentCalls, nestedFactories, standardFactories }),
  };
}

test("Teams Entra object ID is never used as MSAL homeAccountId", () => {
  const app = readFileSync(new URL("../src/Tab/App.tsx", import.meta.url), "utf8");
  const identity = readFileSync(new URL("../src/Tab/copilot.ts", import.meta.url), "utf8");
  assert.doesNotMatch(app, /homeAccountId\s*:\s*context\.user/);
  assert.doesNotMatch(identity, /homeAccountId\s*:/);
});

test("guest cache lookup uses login hint/resource tenant and activates the successfully acquired account", async () => {
  const harness = await identityHarness();
  assert.equal(await harness.identity.acquireApiToken(), "silent-token");
  assert.equal(JSON.stringify(harness.filters[0]), JSON.stringify({ loginHint: "guest@example.com", tenantId }));
  assert.equal(harness.activated.at(-1), guest);
  assert.equal(harness.counts().popups, 0);
  assert.equal(harness.counts().nestedFactories, 1);
  assert.equal(harness.counts().standardFactories, 0);
});

test("network/configuration errors remain visible and never trigger consent popup", async () => {
  const harness = await identityHarness();
  const network = new Error("Network unavailable");
  harness.msal.acquireTokenSilent = async () => { throw network; };
  await assert.rejects(harness.identity.acquireApiToken(), (error) => error === network);
  assert.equal(harness.counts().popups, 0);
});

test("interaction-required errors use popup then activate its account; popup network errors remain unchanged", async () => {
  const harness = await identityHarness();
  harness.msal.acquireTokenSilent = async () => { throw new InteractionRequiredAuthError("Consent required"); };
  assert.equal(await harness.identity.acquireApiToken(), "popup-token");
  assert.equal(harness.activated.at(-1), guest);
  const network = new Error("Popup network/config error");
  harness.msal.acquireTokenPopup = async () => { throw network; };
  await assert.rejects(harness.identity.acquireApiToken(), (error) => error === network);
});

test("tenant mismatch is rejected for Teams context and token result", async () => {
  const harness = await identityHarness();
  await assert.rejects(harness.createIdentity({ tenantId: "other-tenant" }).acquireApiToken(), /Teams tenant does not match/);
  assert.equal(harness.counts().silentCalls, 0);
  harness.msal.acquireTokenSilent = async () => ({
    accessToken: "wrong-tenant", account: { ...guest, tenantId: "other-tenant" },
  });
  await assert.rejects(harness.identity.acquireApiToken(), /outside this app's configured tenant/);
});

test("browser identity retains standard MSAL factory instead of the Teams broker", async () => {
  const harness = await identityHarness(false);
  await harness.identity.acquireApiToken();
  assert.equal(harness.counts().standardFactories, 1);
  assert.equal(harness.counts().nestedFactories, 0);
});
