import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const source = readFileSync(new URL("../src/Tab/App.tsx", import.meta.url), "utf8");
function functionBody(name: string): string {
  const start = source.indexOf(`  async function ${name}(`);
  assert.ok(start >= 0, `${name} exists`);
  const end = source.indexOf("\n  async function ", start + 1);
  return source.slice(start, end < 0 ? source.length : end);
}

test("production initialization exposes no default agent and requires explicit catalog loading", () => {
  assert.match(source, /const initialAgents = config\.demoMode \? config\.demoAgents \|\| \[\] : \[\]/);
  assert.match(source, /Sign in \/ Load catalog/);
  assert.match(source, /disabled=\{busy \|\| \(catalogLoaded && !selectedHarness\.supported\)\}/);
  assert.match(source, /disabled=\{!catalogLoaded \|\| busy \|\| state === "listening"\}/);
});

test("catalog discovery acquires the existing API token, populates the picker and returns ready without invoking agents", () => {
  const body = functionBody("loadCatalog");
  assert.match(body, /loadAgentCatalog\(config, await identity\.acquireApiToken\(\)\)/);
  assert.match(body, /setAgents\(catalog\.agents\)/);
  assert.match(body, /setSelectedAgentId\(selected\?\.id \|\| ""\)/);
  assert.match(body, /setCatalogLoaded\(true\)/);
  assert.match(body, /setState\("ready"\)/);
  assert.doesNotMatch(body, /new (CopilotAgentClient|DemoAgentClient)|\.connect\(|startListening\(|setVoiceConversation\(true\)/);
});

test("connecting before discovery only loads the catalog; after discovery it uses the user's authorized selection", () => {
  const body = functionBody("connect");
  assert.match(body, /if \(!catalogLoaded\) \{\s*await loadCatalog\(\);\s*return;\s*\}/);
  assert.match(body, /const selected = agents\.find\(\(agent\) => agent\.id === selectedAgentId\)/);
  assert.doesNotMatch(body, /config\.defaultAgent|new MicrosoftIdentity|loadAgentCatalog\(/);
  assert.match(body, /new CopilotAgentClient\(identity, selected\)/);
  assert.match(body, /await agentClientRef\.current\.connect\(\)/);
});

test("one-action voice start connects only after discovery and resumes existing conversations without reconnecting", () => {
  const body = functionBody("startVoiceConversation");
  assert.match(body, /if \(!catalogLoaded\) \{\s*await loadCatalog\(\);\s*return;\s*\}/);
  assert.match(body, /setVoiceConversation\(true\);\s*if \(!agentClientRef\.current\) \{\s*await connect\(\);\s*return;\s*\}\s*await startListening\(true\)/);
  assert.match(source, /disabled=\{!catalogLoaded \|\| !selectedHarness\.supported \|\| busy \|\| state === "complete"\}/);
  assert.match(source, /const connected = Boolean\(agentClientRef\.current\)/);
});

test("Connect preserves explicitly selected input language while deliberate agent changes default to output Locale", () => {
  const connect = functionBody("connect");
  assert.doesNotMatch(connect, /setInputLanguage\(selected\.locale\)/);
  assert.match(connect, /await speechRef\.current\?\.setAgent\(selected\);\s*speechRef\.current\?\.setInputLanguage\(inputLanguage\)/);
  const select = functionBody("selectAgent");
  assert.match(select, /setInputLanguage\(selected\.locale\)/);
  assert.match(select, /speechRef\.current\?\.setInputLanguage\(selected\.locale\)/);
});
