import type {
  AgentDefinition,
  AppConfig,
  CatalogResult,
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

function requiredBoolean(value: unknown, field: string, itemId: string): boolean {
  if (value === true || value === 1 || value === "1" || value === "true") {
    return true;
  }
  if (value === false || value === 0 || value === "0" || value === "false") {
    return false;
  }
  throw new Error(`Catalog item ${itemId} has an invalid ${field}.`);
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
  const audienceValue = requiredString(record.Audience, "Audience", itemId).toLowerCase();
  if (audienceValue !== "everyone" && audienceValue !== "restricted") {
    throw new Error(`Catalog item ${itemId} has an invalid Audience.`);
  }

  const endsConversation = requiredBoolean(record.EndsConversation, "EndsConversation", itemId);
  const displayName = requiredString(record.Title, "Title", itemId);
  return {
    id: itemId,
    displayName,
    description: optionalString(record.Description),
    environmentId,
    schemaName,
    completionPhrase: endsConversation
      ? requiredString(record.CompletionPhrase, "CompletionPhrase", itemId)
      : optionalString(record.CompletionPhrase) || "",
    endsConversation,
    welcomeMessage: optionalString(record.WelcomeMessage) || `Ready to chat with ${displayName}.`,
    locale: requiredString(record.Locale, "Locale", itemId),
    voiceName: requiredString(record.VoiceName, "VoiceName", itemId),
    avatarCharacter: requiredString(
      record.AvatarCharacter,
      "AvatarCharacter",
      itemId,
    ),
    avatarStyle: requiredString(record.AvatarStyle, "AvatarStyle", itemId),
    audience: audienceValue === "restricted"
      ? "Restricted"
      : "Everyone",
    harness: optionalString(record.Harness),
  };
}

export async function loadAgentCatalog(
  config: AppConfig,
  accessToken: string,
): Promise<CatalogResult> {
  if (!config.catalogEnabled) {
    throw new CatalogAccessError("The SharePoint agent catalog is not configured.");
  }

  const response = await fetch("/api/catalog", {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
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

  const agents = result.agents
    .map((record, index) => mapAgent(record, index))
    .filter((agent): agent is AgentDefinition => Boolean(agent))
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
