import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

test("Entra scope bootstrap and full manifest use the same delegated permission", () => {
  const bootstrap = JSON.parse(fs.readFileSync("aad.api.manifest.json", "utf8"));
  const full = JSON.parse(fs.readFileSync("aad.manifest.json", "utf8"));
  assert.deepEqual(bootstrap.api.oauth2PermissionScopes, full.api.oauth2PermissionScopes);
  const scope = full.api.oauth2PermissionScopes[0];
  assert.equal(scope.value, "access_as_user");
  assert.equal(scope.type, "User");
  for (const client of full.api.preAuthorizedApplications) {
    assert.deepEqual(client.delegatedPermissionIds, [scope.id]);
  }
});

test("registration preserves the draft approval gate and restricts lookup to the configured environment", () => {
  const template = JSON.parse(fs.readFileSync("catalog/agent-registration-flow.template.json", "utf8"));
  const resolve = template.properties.definition.actions.Resolve_agent;
  const condition = resolve.actions.Validate_URL;
  assert.match(condition.expression, /__ENVIRONMENT_ID__/);
  assert.match(condition.expression, /copilotstudio\.microsoft\.com/);
  assert.match(condition.expression, /uriHost/);
  const draft = condition.actions.Fill_catalog_draft.inputs.parameters;
  assert.equal(draft["item/__FIELD_ENABLED__"], "FALSE");
  assert.equal(draft["item/EndsConversation"], false);
  assert.match(draft["item/Harness"], /GenerativeAIRecognizer/);
  const failure = template.properties.definition.actions.Report_lookup_failure;
  assert.equal(failure.inputs.parameters["item/__FIELD_ENABLED__"], "FALSE");
  assert.match(failure.inputs.parameters["item/RegistrationStatus"], /Nothing was enabled/);
});

test("catalog response run history protects private inputs", () => {
  const template = JSON.parse(fs.readFileSync("catalog/catalog-flow.template.json", "utf8"));
  const actions = template.properties.definition.actions;
  for (const name of ["Read_catalog", "Select_catalog", "Respond"]) {
    assert.ok(actions[name].runtimeConfiguration.secureData.properties.includes("inputs"));
  }
  assert.deepEqual(actions.Respond.runtimeConfiguration.secureData.properties, ["inputs"]);
});

test("MSAL v5 callback uses a dedicated redirect bridge instead of a placeholder", () => {
  const entry = fs.readFileSync("src/Tab/auth-callback.ts", "utf8");
  const page = fs.readFileSync("auth-callback.html", "utf8");
  const host = fs.readFileSync("src/index.ts", "utf8");
  assert.match(entry, /@azure\/msal-browser\/redirect-bridge/);
  assert.match(entry, /broadcastResponseToMainFrame\(\)/);
  assert.match(page, /auth-callback\.ts/);
  assert.match(host, /sendFile\(path\.join\(__dirname, "client", "auth-callback.html"\)\)/);
  assert.match(host, /frame-src 'self' https:\/\/login\.microsoftonline\.com/);
  assert.match(host, /script-src 'self'; worker-src 'self' blob: data:/);
  assert.doesNotMatch(host, /setHeader\("Cross-Origin-Opener-Policy"/);
});

test("public config exposes operator-editable cost notices with explicit default rates", () => {
  const host = fs.readFileSync("src/index.ts", "utf8");
  assert.match(host, /VOICE_COST_NOTICE\?\.trim\(\)/);
  assert.match(host, /AVATAR_COST_NOTICE\?\.trim\(\)/);
  assert.match(host, /\$1\/hour[\s\S]*?\$15\/million characters/);
  assert.match(host, /\$0\.50\/min while connected, including silence/);
});
