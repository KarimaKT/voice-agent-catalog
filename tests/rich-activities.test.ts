import assert from "node:assert/strict";
import test from "node:test";
import { Activity, ActivityTypes } from "@microsoft/agents-activity";

import {
  collectAgentTurn,
  CopilotAgentClient,
  createMessageActivity,
  normalizeActivity,
  toSpokenText,
  type MicrosoftIdentity,
} from "../src/Tab/copilot";
import type { AgentDefinition } from "../src/Tab/types";

function activity(json: Record<string, unknown>): Activity {
  return Object.assign(new Activity(String(json.type)), json);
}

async function* replies(...values: Activity[]): AsyncIterable<Activity> {
  yield* values;
}

const card = {
  type: "AdaptiveCard",
  version: "1.5",
  body: [{
    type: "Input.Text",
    id: "customer",
    label: "Customer",
    isRequired: true,
    errorMessage: "Enter a customer.",
  }],
  actions: [{ type: "Action.Submit", title: "Save", data: { command: "save", orderId: 42 } }],
};

test("real attachment-only activity preserves the complete card, safe images and files", () => {
  const message = normalizeActivity(activity({
    type: "message",
    conversation: { id: "conversation-1" },
    summary: "Do not narrate this card and its body",
    attachments: [
      { contentType: "application/vnd.microsoft.card.adaptive", content: card },
      { contentType: "image/png", contentUrl: "https://example.com/image.png", name: "Product" },
      { contentType: "application/pdf", contentUrl: "https://example.com/report.pdf", name: "Report" },
    ],
  }));
  assert.equal(message.text, "");
  assert.equal(message.speak, undefined);
  assert.equal(message.attachments?.length, 3);
  assert.deepEqual(message.attachments?.[0], { kind: "adaptiveCard", name: undefined, content: card });
  assert.equal(message.attachments?.[1].kind, "image");
  assert.equal(message.attachments?.[2].kind, "file");
});

test("unsafe and unknown attachments have visible fallback rather than disappearing", () => {
  const message = normalizeActivity(activity({
    type: "message",
    attachments: [
      { contentType: "image/png", contentUrl: "javascript:alert(1)" },
      { contentType: "application/vnd.microsoft.card.hero", content: { title: "Unsupported" } },
      { contentType: "application/vnd.microsoft.card.adaptive" },
    ],
  }));
  assert.equal(message.attachments?.length, 3);
  assert.ok(message.attachments?.every((attachment) => attachment.kind === "unsupported"));
});

test("citations and known suggested actions survive while unsafe/unsupported actions are explained", () => {
  const message = normalizeActivity(activity({
    type: "message",
    text: "Sources",
    entities: [{
      type: "https://schema.org/Message",
      citation: [
        { appearance: { name: "Safe source", abstract: "Description", url: "https://example.com/source" } },
        { appearance: { name: "Unsafe source", url: "data:text/html,unsafe" } },
      ],
    }],
    suggestedActions: { actions: [
      { type: "imBack", title: "Continue", value: "continue" },
      { type: "postBack", title: "Choose", value: { orderId: 42 } },
      { type: "openUrl", title: "Safe", value: "https://example.com/" },
      { type: "openUrl", title: "Unsafe", value: "javascript:alert(1)" },
      { type: "invoke", title: "Unknown", value: "ignored" },
    ] },
  }));
  assert.equal(message.citations?.[0].url, "https://example.com/source");
  assert.equal(message.citations?.[1].url, undefined);
  assert.deepEqual(message.suggestedActions?.[0], { title: "Continue", kind: "message", value: "continue" });
  assert.deepEqual(message.suggestedActions?.[1]?.value, { orderId: 42 });
  assert.equal(message.suggestedActions?.[1]?.kind, "submit");
  assert.equal(message.suggestedActions?.[2]?.url, "https://example.com/");
  assert.match(message.suggestedActions?.[3]?.unsupported || "", /unsafe/);
  assert.match(message.suggestedActions?.[4]?.unsupported || "", /Unsupported/);
});

test("Submit objects, including merged input values, stay in Activity.value in the same conversation", () => {
  const value = { command: "save", orderId: 42, customer: "Contoso", approved: true };
  const outgoing = createMessageActivity({ title: "Save", kind: "submit", value }, "continuous-1", "fr-FR");
  assert.equal(outgoing.type, ActivityTypes.Message);
  assert.equal(outgoing.text, undefined);
  assert.deepEqual(outgoing.value, value);
  assert.equal(outgoing.conversation?.id, "continuous-1");
  assert.equal(outgoing.locale, "fr-FR");
  assert.deepEqual(createMessageActivity({ title: "Submit", kind: "submit" }, "continuous-1").value, {});
  assert.equal(createMessageActivity("Hello", "continuous-1").text, "Hello");
});

test("streams accept attachment-only and speak-only turns and honor EndOfConversation without text", async () => {
  const turn = await collectAgentTurn(replies(
    activity({ type: "typing", conversation: { id: "continuous-1" } }),
    activity({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] }),
    activity({ type: "message", speak: "A short spoken reply." }),
    activity({ type: "endOfConversation", code: "completedSuccessfully" }),
  ));
  assert.equal(turn.conversationId, "continuous-1");
  assert.equal(turn.messages.length, 2);
  assert.equal(turn.messages[0].text, "");
  assert.equal(turn.completed, true);
  const next = await collectAgentTurn(replies(activity({ type: "message", text: "Continue" })), turn.conversationId);
  assert.equal(next.conversationId, "continuous-1");
  assert.equal(next.completed, false);
  const ended = await collectAgentTurn(replies(
    activity({ type: "endOfConversation", code: "completedSuccessfully" }),
  ), turn.conversationId);
  assert.equal(ended.messages.length, 0);
  assert.equal(ended.completed, true);
  assert.equal(ended.conversationId, "continuous-1");
});

test("silent startup preserves the SDK's conversation metadata even with no activities", async () => {
  const response = { activities: [], conversationId: "header-conversation" };
  const turn = await collectAgentTurn(response.activities, response.conversationId);
  assert.equal(turn.conversationId, "header-conversation");
  assert.equal(turn.messages.length, 0);
  assert.equal(turn.completed, false);
});

test("Copilot transport sends card submissions through existing streaming conversation", async () => {
  const definition = {
    displayName: "Agent", environmentId: "environment-1", schemaName: "agent",
    harness: "Standard", locale: "de-DE",
  } as AgentDefinition;
  const client = new CopilotAgentClient({} as MicrosoftIdentity, definition);
  const outgoing: Activity[] = [];
  const transport = {
    sendActivityStreaming(value: Activity, id: string) {
      assert.equal(id, "continuous-1");
      outgoing.push(value);
      return replies(activity({ type: "message", text: "Saved", conversation: { id } }));
    },
  };
  Object.assign(client, { client: transport, conversationId: "continuous-1" });
  await client.send({ title: "Save", kind: "submit", value: { customer: "Contoso" } });
  await client.send("What next?");
  assert.deepEqual(outgoing[0].value, { customer: "Contoso" });
  assert.equal(outgoing[1].text, "What next?");
  assert.equal(outgoing[0].conversation?.id, outgoing[1].conversation?.id);
});

test("speech text never derives content from attachments", () => {
  const message = normalizeActivity(activity({
    type: "message", text: "Here is the form.",
    attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }],
  }));
  assert.equal(toSpokenText(message.text), "Here is the form.");
  assert.doesNotMatch(toSpokenText(message.text), /Customer|Save|orderId/);
});
