import assert from "node:assert/strict";
import test from "node:test";

import { validateAndAuthorizeCatalog } from "../src/catalog-policy";

const base = {
  Id: "7",
  Title: "Pat",
  EnvironmentId: "11111111-1111-1111-1111-111111111111",
  SchemaName: "cr_test",
  Enabled: true,
  CompletionPhrase: "Done.",
  EndsConversation: true,
  Locale: "en-US",
  VoiceName: "en-US-AvaMultilingualNeural",
  AvatarCharacter: "lisa",
  AvatarStyle: "casual",
  Audience: "Everyone",
  Harness: "standard",
};

test("authorizes Everyone and strips AllowedUsers", () => {
  const [agent] = validateAndAuthorizeCatalog(
    [{ ...base, AllowedUsers: [{ Email: "secret@example.com" }] }],
    { oid: "oid", email: "user@example.com" },
  );
  assert.equal(agent.Title, "Pat");
  assert.equal("AllowedUsers" in agent, false);
});

test("Restricted uses exact normalized email equality", () => {
  const record = {
    ...base,
    Audience: "Restricted",
    AllowedUsers: [{ Email: " Allowed@Example.com " }],
  };
  assert.equal(
    validateAndAuthorizeCatalog([record], { oid: "oid", email: "allowed@example.com" }).length,
    1,
  );
  assert.equal(
    validateAndAuthorizeCatalog([record], { oid: "oid", email: "allowed+other@example.com" }).length,
    0,
  );
});

test("disabled rows are not substituted and CSV boolean values are accepted", () => {
  assert.deepEqual(validateAndAuthorizeCatalog([{ Enabled: false }], { oid: "oid" }), []);
  assert.equal(validateAndAuthorizeCatalog([
    { ...base, Enabled: "TRUE", EndsConversation: "FALSE" },
  ], { oid: "oid" })[0].EndsConversation, false);
  assert.throws(() => validateAndAuthorizeCatalog([{ ...base, Enabled: "perhaps" }], { oid: "oid" }));
});

test("conversational agents do not require a terminal phrase", () => {
  const [agent] = validateAndAuthorizeCatalog([
    { ...base, EndsConversation: false, CompletionPhrase: "" },
  ], { oid: "oid" });
  assert.equal(agent.CompletionPhrase, "");
});

test("rejects unknown Audience, missing completion phrase, and invalid restrictions", () => {
  assert.throws(
    () => validateAndAuthorizeCatalog([{ ...base, Audience: "Partners" }], { oid: "oid" }),
    /invalid Audience/,
  );
  assert.throws(
    () => validateAndAuthorizeCatalog([{ ...base, CompletionPhrase: " " }], { oid: "oid" }),
    /no CompletionPhrase/,
  );
  assert.throws(
    () => validateAndAuthorizeCatalog(
      [{ ...base, Audience: "Restricted", AllowedUsers: ["not-an-email"] }],
      { oid: "oid" },
    ),
    /no valid AllowedUsers/,
  );
});
