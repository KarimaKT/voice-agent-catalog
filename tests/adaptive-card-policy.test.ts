import assert from "node:assert/strict";
import test from "node:test";

import { safeHttpUrl, sanitizeAdaptiveCard } from "../src/Tab/adaptive-card-policy";

test("only absolute HTTP(S) URLs without credentials or control characters are allowed", () => {
  assert.equal(safeHttpUrl("https://example.com/image.png"), "https://example.com/image.png");
  assert.equal(safeHttpUrl("http://example.com/"), "http://example.com/");
  assert.equal(safeHttpUrl("https://example.com/My image.png"), "https://example.com/My%20image.png");
  for (const value of [
    "javascript:alert(1)", "data:image/svg+xml,unsafe", "file:///C:/secret",
    "//example.com/image.png", "/image.png", "https://user:password@example.com/",
    "https://example.com/\nunsafe", "blob:https://example.com/image", undefined, {},
  ]) {
    assert.equal(safeHttpUrl(value), undefined);
  }
});

test("the official renderer receives unchanged input/validation, Submit data, safe imagery and ShowCards", () => {
  const data = { verb: "save", order: { id: 42 }, type: "Action.Execute", url: "opaque business data" };
  const card = {
    type: "AdaptiveCard", version: "1.6",
    body: [
      { type: "Input.Text", id: "email", label: "Email", isRequired: true, regex: "^.+@.+$", errorMessage: "Email required" },
      { type: "Input.ChoiceSet", id: "priority", choices: [{ title: "High", value: "high" }] },
      { type: "Image", url: "https://example.com/image.png" },
    ],
    actions: [
      { type: "Action.Submit", title: "Save", data, associatedInputs: "auto" },
      { type: "Action.OpenUrl", title: "Read", url: "https://example.com/" },
      { type: "Action.ShowCard", title: "More", card: { type: "AdaptiveCard", body: [{ type: "TextBlock", text: "More" }] } },
    ],
  };
  const sanitized = sanitizeAdaptiveCard(card);
  assert.deepEqual(sanitized.card, card);
  assert.deepEqual(sanitized.warnings, []);
});

test("unsafe resource URLs and actions are removed recursively without mutating original card", () => {
  const card = {
    type: "AdaptiveCard", version: "1.5",
    backgroundImage: { url: "data:image/svg+xml,unsafe" },
    body: [{
      type: "Container", items: [
        { type: "Image", url: "javascript:alert(1)" },
        { type: "Input.Text", id: "value", inlineAction: { type: "Action.OpenUrl", title: "Bad", url: "file:///C:/secret" } },
        { type: "TextBlock", text: "Safe", selectAction: { type: "Action.Execute", title: "Invoke" } },
        { type: "ImageSet", images: [{ url: "javascript:alert(1)" }] },
      ],
    }],
    actions: [{
      type: "Action.ShowCard", title: "More", iconUrl: "data:image/svg+xml,unsafe",
      card: { type: "AdaptiveCard", body: [{ type: "Image", url: "blob:https://example.com/id" }] },
    }],
    refresh: { action: { type: "Action.Execute", verb: "refresh" } },
    authentication: { tokenExchangeResource: {} },
  };
  const before = JSON.stringify(card);
  const result = sanitizeAdaptiveCard(card);
  assert.ok(result.card);
  assert.ok(result.warnings.length >= 6);
  assert.doesNotMatch(JSON.stringify(result.card), /javascript:|file:|blob:|data:image|Action.Execute/);
  assert.equal(JSON.stringify(card), before);
});

test("unsupported elements/actions surface notices and safe authored fallbacks", () => {
  const result = sanitizeAdaptiveCard({
    type: "AdaptiveCard", version: "1.5",
    body: [{ type: "Input.File" }, { type: "Unknown", fallback: { type: "TextBlock", text: "Use the text response" } }],
    actions: [
      { type: "Action.Execute", title: "Not supported", fallback: { type: "Action.Submit", title: "Submit", data: { command: "save" } } },
      { type: "Action.Unknown", title: "Unknown" },
    ],
  });
  assert.equal((result.card?.body as Record<string, unknown>[])[0].type, "TextBlock");
  assert.equal((result.card?.body as Record<string, unknown>[])[1].text, "Use the text response");
  assert.equal((result.card?.actions as Record<string, unknown>[]).length, 1);
  assert.match(result.warnings.join(" "), /Input.File.*Unknown.*Action.Execute.*Action.Unknown/);
});

test("invalid cards, future versions, excessive size and depth have explicit bounded fallback", () => {
  let nested: unknown = { type: "TextBlock", text: "Deep" };
  for (let index = 0; index < 30; index++) nested = { type: "Container", items: [nested] };
  const base = { type: "AdaptiveCard", version: "1.5" };
  for (const value of [
    null, { ...base, type: "HeroCard" }, { ...base, version: "2.0" }, { ...base, body: "invalid" },
    { ...base, version: "1.7" }, { ...base, body: [nested] },
    { ...base, body: [{ type: "TextBlock", text: "x".repeat(128_000) }] },
  ]) {
    const result = sanitizeAdaptiveCard(value);
    assert.equal(result.card, undefined);
    assert.ok(result.warnings.length);
    assert.match(result.warnings[0], /Use text to respond/);
  }
});

test("action count is bounded with a visible notice instead of SDK silently dropping extras", () => {
  const result = sanitizeAdaptiveCard({
    type: "AdaptiveCard", version: "1.5",
    actions: Array.from({ length: 21 }, (_, index) => ({ type: "Action.Submit", title: `${index}` })),
  });
  assert.equal((result.card?.actions as unknown[]).length, 20);
  assert.match(result.warnings.join(" "), /first 20/);
});
