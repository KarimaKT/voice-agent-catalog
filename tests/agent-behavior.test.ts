import assert from "node:assert/strict";
import test from "node:test";

import { completesConversation, getHarnessSupport, voiceInterfaceInstructions } from "../src/Tab/agent-behavior";

test("completion requires EndsConversation and exact normalized equality", () => {
  const agent = { endsConversation: true, completionPhrase: "Thank you. Done." };
  assert.equal(completesConversation(agent, "  THANK you.   done.  "), true);
  assert.equal(completesConversation(agent, "Thank you. Done. Additional text"), false);
  assert.equal(completesConversation(agent, '"Thank you. Done."'), false);
  assert.equal(completesConversation(agent, "Not yet: Thank you. Done."), false);
  assert.equal(completesConversation({ endsConversation: true, completionPhrase: "" }, ""), false);
  assert.equal(
    completesConversation({ ...agent, endsConversation: false }, "Thank you. Done."),
    false,
  );
});

test("voice context honors the output locale and preserves agent authorization", () => {
  const instruction = voiceInterfaceInstructions("fr-FR");
  assert.match(instruction, /answer in fr-FR/);
  assert.match(instruction, /existing instructions and authorization rules/);
  assert.match(instruction, /not a request to execute a tool/);
});

test("only Standard harness transport is supported", () => {
  assert.deepEqual(getHarnessSupport(" STANDARD "), {
    supported: true,
    label: "Standard",
  });
  assert.equal(getHarnessSupport("GitHub Copilot").supported, false);
  assert.equal(getHarnessSupport(undefined).supported, false);
});
