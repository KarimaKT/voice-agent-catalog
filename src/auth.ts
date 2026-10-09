import { createRemoteJWKSet, jwtVerify, type JWTVerifyGetKey, type JWTPayload } from "jose";
import type { Request } from "express";

export interface VerifiedUser {
  oid: string;
  email?: string;
}

const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const keySets = new Map<string, ReturnType<typeof createRemoteJWKSet>>();

export function normalizeUserEmail(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const email = value.trim().toLowerCase();
  return emailPattern.test(email) ? email : undefined;
}

export function userFromVerifiedClaims(
  claims: JWTPayload,
  tenantId: string,
): VerifiedUser {
  if (claims.tid !== tenantId) {
    throw new Error("The access token tenant is invalid.");
  }
  if (typeof claims.oid !== "string" || !claims.oid.trim()) {
    throw new Error("The access token has no object identifier.");
  }
  return {
    oid: claims.oid.trim().toLowerCase(),
    email:
      normalizeUserEmail(claims.preferred_username) ||
      normalizeUserEmail(claims.upn) ||
      normalizeUserEmail(claims.email),
  };
}

function bearerToken(request: Request): string {
  const authorization = request.header("Authorization");
  const match = authorization?.match(/^Bearer\s+(\S+)$/i);
  if (!match) {
    throw new Error("A bearer access token is required.");
  }
  return match[1];
}

export async function verifyApiToken(
  token: string,
  tenantId: string,
  audience: string,
  keySet: JWTVerifyGetKey,
): Promise<VerifiedUser> {
  const { payload } = await jwtVerify(token, keySet, {
    issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
    audience,
    algorithms: ["RS256"],
    requiredClaims: ["exp", "tid", "oid", "scp"],
  });
  const scopes = typeof payload.scp === "string" ? payload.scp.split(/\s+/) : [];
  if (!scopes.includes("access_as_user")) {
    throw new Error("The access token is missing the access_as_user scope.");
  }
  return userFromVerifiedClaims(payload, tenantId);
}

export async function authenticateApiRequest(
  request: Request,
  tenantId: string,
  audience: string,
): Promise<VerifiedUser> {
  const token = bearerToken(request);
  let keySet = keySets.get(tenantId);
  if (!keySet) {
    keySet = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`),
    );
    keySets.set(tenantId, keySet);
  }

  return verifyApiToken(token, tenantId, audience, keySet);
}
