import type { AgentDefinition, AppConfig } from "./types";

type CatalogAgentRecord = Record<string, unknown>;

interface CatalogResponse {
  agents?: CatalogAgentRecord[];
}

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
  };
}

export async function loadAgentCatalog(config: AppConfig): Promise<AgentDefinition[]> {
  if (!config.catalogEnabled) {
    return [config.defaultAgent];
  }

  const response = await fetch("/api/catalog", {
    headers: { Accept: "application/json" },
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
    .map((record, index) => mapAgent(record, index, config.defaultAgent))
    .filter((agent): agent is AgentDefinition => Boolean(agent))
    .sort((left, right) => left.displayName.localeCompare(right.displayName));

  if (agents.length === 0) {
    throw new Error("Catalog flow returned no enabled agents.");
  }
  return agents;
}
