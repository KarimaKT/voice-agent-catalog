import type {
  AgentDefinition,
  AppConfig,
  CatalogResult,
  CatalogUser,
} from "./types";

type CatalogAgentRecord = Record<string, unknown>;

interface CatalogResponse {
  agents?: CatalogAgentRecord[];
  canManageCatalog?: unknown;
  configurationUrl?: unknown;
}

export class CatalogAccessError extends Error {}

const environmentIdPattern =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const schemaNamePattern = /^[A-Za-z][A-Za-z0-9_.-]{0,199}$/;

function requiredString(value: unknown, field: string, itemId: string): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`Catalog item ${itemId} has no ${field}.`);
  }
  return value.trim();
}

function optionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function isEnabled(value: unknown): boolean {
  return value === true || value === 1 || value === "1" || value === "true";
}

function normalizeEmail(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const email = value.trim().toLocaleLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : undefined;
}

function normalizeUsers(value: unknown): CatalogUser[] {
  const source = Array.isArray(value) ? value : value ? [value] : [];
  return source.flatMap((entry) => {
    if (typeof entry === "string") {
      const email = normalizeEmail(entry);
      return email ? [{ email }] : [];
    }
    if (!entry || typeof entry !== "object") {
      return [];
    }
    const user = entry as Record<string, unknown>;
    const email =
      normalizeEmail(user.Email) ||
      normalizeEmail(user.email) ||
      normalizeEmail(user.UserPrincipalName) ||
      normalizeEmail(user.userPrincipalName);
    if (!email) {
      return [];
    }
    return [{
      email,
      displayName:
        optionalString(user.DisplayName) ||
        optionalString(user.displayName) ||
        optionalString(user.Title),
    }];
  });
}

function safeHttpUrl(value: unknown): string | undefined {
  const text = optionalString(value);
  if (!text) {
    return undefined;
  }
  try {
    const url = new URL(text);
    return url.protocol === "https:" || url.protocol === "http:" ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function mapAgent(
  record: CatalogAgentRecord,
  index: number,
  defaults: AgentDefinition,
): AgentDefinition | undefined {
  const itemId = optionalString(record.Id) || optionalString(record.id) || String(index + 1);
  if (!isEnabled(record.Enabled)) {
    return undefined;
  }

  const environmentId = requiredString(record.EnvironmentId, "EnvironmentId", itemId);
  const schemaName = requiredString(record.SchemaName, "SchemaName", itemId);
  if (!environmentIdPattern.test(environmentId)) {
    throw new Error(`Catalog item ${itemId} has an invalid EnvironmentId.`);
  }
  if (!schemaNamePattern.test(schemaName)) {
    throw new Error(`Catalog item ${itemId} has an invalid SchemaName.`);
  }

  return {
    id: itemId,
    displayName: requiredString(record.Title, "Title", itemId),
    description: optionalString(record.Description),
    environmentId,
    schemaName,
    completionPhrase:
      optionalString(record.CompletionPhrase) || defaults.completionPhrase,
    welcomeMessage: optionalString(record.WelcomeMessage) || defaults.welcomeMessage,
    locale: requiredString(record.Locale, "Locale", itemId),
    voiceName: requiredString(record.VoiceName, "VoiceName", itemId),
    avatarCharacter: requiredString(
      record.AvatarCharacter,
      "AvatarCharacter",
      itemId,
    ),
    avatarStyle: requiredString(record.AvatarStyle, "AvatarStyle", itemId),
    audience:
      optionalString(record.Audience)?.toLocaleLowerCase() === "restricted"
        ? "Restricted"
        : "Everyone",
    allowedUsers: normalizeUsers(record.AllowedUsers),
    harness: optionalString(record.Harness),
  };
}

export async function loadAgentCatalog(
  config: AppConfig,
  userEmail?: string,
): Promise<CatalogResult> {
  if (!config.catalogEnabled) {
    return { agents: [config.defaultAgent], canManageCatalog: false };
  }

  const response = await fetch("/api/catalog", {
    headers: {
      Accept: "application/json",
      ...(userEmail ? { "X-Voice-Catalog-User": userEmail } : {}),
    },
  });
  if (!response.ok) {
    const body = (await response.json().catch(() => undefined)) as
      | { error?: string }
      | undefined;
    throw new Error(body?.error || `Catalog request failed (${response.status}).`);
  }

  const result = (await response.json()) as CatalogResponse;
  if (!Array.isArray(result.agents)) {
    throw new Error("Catalog flow returned an invalid response.");
  }

  const normalizedEmail = normalizeEmail(userEmail);
  const agents = result.agents
    .map((record, index) => mapAgent(record, index, config.defaultAgent))
    .filter((agent): agent is AgentDefinition => Boolean(agent))
    .filter(
      (agent) =>
        agent.audience !== "Restricted" ||
        Boolean(
          normalizedEmail &&
            agent.allowedUsers?.some((user) => user.email === normalizedEmail),
        ),
    )
    .sort((left, right) => left.displayName.localeCompare(right.displayName));

  if (agents.length === 0) {
    throw new CatalogAccessError("No catalog agents are available for this user.");
  }
  const canManageCatalog = result.canManageCatalog === true;
  return {
    agents,
    canManageCatalog,
    configurationUrl: canManageCatalog
      ? safeHttpUrl(result.configurationUrl)
      : undefined,
  };
}
