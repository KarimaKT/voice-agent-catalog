import assert from "node:assert/strict";
import test from "node:test";
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from "jose";

import { normalizeUserEmail, verifyApiToken } from "../src/auth";

const tenant = "11111111-1111-1111-1111-111111111111";
const audience = "22222222-2222-2222-2222-222222222222";
const issuer = `https://login.microsoftonline.com/${tenant}/v2.0`;

test("API tokens require a valid signature, issuer, audience, expiry and delegated scope", async () => {
  const keys = await generateKeyPair("RS256");
  const jwk = await exportJWK(keys.publicKey);
  const resolveKey = createLocalJWKSet({ keys: [{ ...jwk, kid: "test", alg: "RS256" }] });
  async function token(overrides: Record<string, unknown> = {}, expires = "2m") {
    return new SignJWT({
      tid: tenant, oid: "33333333-3333-3333-3333-333333333333",
      upn: " User@Example.com ", scp: "access_as_user", ...overrides,
    }).setProtectedHeader({ alg: "RS256", kid: "test" })
      .setIssuer(issuer).setAudience(audience).setIssuedAt().setExpirationTime(expires)
      .sign(keys.privateKey);
  }
  assert.deepEqual(await verifyApiToken(await token(), tenant, audience, resolveKey), {
    oid: "33333333-3333-3333-3333-333333333333", email: "user@example.com",
  });
  for (const badClaims of [
    { tid: "other" }, { oid: "" }, { scp: "" }, { scp: "access_as_user_extra" },
  ]) {
    await assert.rejects(verifyApiToken(await token(badClaims), tenant, audience, resolveKey));
  }
  await assert.rejects(verifyApiToken(await token(), "other-tenant", audience, resolveKey));
  await assert.rejects(verifyApiToken(await token(), tenant, "other-api", resolveKey));
  await assert.rejects(verifyApiToken(await token({}, "-1m"), tenant, audience, resolveKey));
  const otherKeys = await generateKeyPair("RS256");
  const forged = await new SignJWT({ tid: tenant, oid: "oid", scp: "access_as_user" })
    .setProtectedHeader({ alg: "RS256", kid: "test" }).setIssuer(issuer)
    .setAudience(audience).setExpirationTime("2m").sign(otherKeys.privateKey);
  await assert.rejects(verifyApiToken(forged, tenant, audience, resolveKey));
});

test("email normalization rejects invalid identity hints", () => {
  assert.equal(normalizeUserEmail(" Person@Example.com "), "person@example.com");
  assert.equal(normalizeUserEmail("not-an-email"), undefined);
  assert.equal(normalizeUserEmail(undefined), undefined);
});
