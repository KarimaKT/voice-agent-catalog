import { MicrosoftIdentity } from "./copilot";
import type { AgentDefinition, AppConfig } from "./types";

interface GraphSite {
  id?: string;
}

interface GraphList {
  id?: string;
  displayName?: string;
}

interface GraphListResponse {
  value?: GraphList[];
}

type AgentFields = Record<string, unknown>;

interface GraphColumn {
  name?: string;
  displayName?: string;
}

interface GraphColumnsResponse {
  value?: GraphColumn[];
}

interface AgentColumnNames {
  Title: string;
  Description: string;
  EnvironmentId: string;
  SchemaName: string;
  Enabled: string;
  CompletionPhrase: string;
  WelcomeMessage: string;
  Locale: string;
  VoiceName: string;
  AvatarCharacter: string;
  AvatarStyle: string;
}

interface GraphListItem {
  id?: string;
  fields?: AgentFields;
}

interface GraphItemsResponse {
  value?: GraphListItem[];
  "@odata.nextLink"?: string;
}

const graphRoot = "https://graph.microsoft.com/v1.0";
const graphScopes = ["https://graph.microsoft.com/Sites.Read.All"];
const environmentIdPattern =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const schemaNamePattern = /^[A-Za-z][A-Za-z0-9_.-]{0,199}$/;

async function graphGet<T>(url: string, token: string): Promise<T> {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const requestId = response.headers.get("request-id");
    throw new Error(
      `Microsoft Graph request failed (${response.status})${
        requestId ? `, request ${requestId}` : ""
      }. Check SharePoint access and Graph consent.`,
    );
  }

  return (await response.json()) as T;
}

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

function resolveColumns(columns: GraphColumn[]): AgentColumnNames {
  const byDisplayName = new Map(
    columns
      .filter((column) => column.name && column.displayName)
      .map((column) => [column.displayName, column.name] as const),
  );
  const required = [
    "Title",
    "Description",
    "EnvironmentId",
    "SchemaName",
    "Enabled",
    "CompletionPhrase",
    "WelcomeMessage",
    "Locale",
    "VoiceName",
    "AvatarCharacter",
    "AvatarStyle",
  ] as const;
  const missing = required.filter((name) => !byDisplayName.get(name));
  if (missing.length > 0) {
    throw new Error(`Agent catalog is missing columns: ${missing.join(", ")}.`);
  }

  return {
    Title: byDisplayName.get("Title")!,
    Description: byDisplayName.get("Description")!,
    EnvironmentId: byDisplayName.get("EnvironmentId")!,
    SchemaName: byDisplayName.get("SchemaName")!,
    Enabled: byDisplayName.get("Enabled")!,
    CompletionPhrase: byDisplayName.get("CompletionPhrase")!,
    WelcomeMessage: byDisplayName.get("WelcomeMessage")!,
    Locale: byDisplayName.get("Locale")!,
    VoiceName: byDisplayName.get("VoiceName")!,
    AvatarCharacter: byDisplayName.get("AvatarCharacter")!,
    AvatarStyle: byDisplayName.get("AvatarStyle")!,
  };
}

function mapAgent(
  item: GraphListItem,
  columns: AgentColumnNames,
  defaults: AgentDefinition,
): AgentDefinition | undefined {
  const itemId = item.id || "unknown";
  const fields = item.fields;
  if (!fields || !isEnabled(fields[columns.Enabled])) {
    return undefined;
  }

  const environmentId = requiredString(
    fields[columns.EnvironmentId],
    "EnvironmentId",
    itemId,
  );
  const schemaName = requiredString(fields[columns.SchemaName], "SchemaName", itemId);
  if (!environmentIdPattern.test(environmentId)) {
    throw new Error(`Catalog item ${itemId} has an invalid EnvironmentId.`);
  }
  if (!schemaNamePattern.test(schemaName)) {
    throw new Error(`Catalog item ${itemId} has an invalid SchemaName.`);
  }

  return {
    id: itemId,
    displayName: requiredString(fields[columns.Title], "Title", itemId),
    description: optionalString(fields[columns.Description]),
    environmentId,
    schemaName,
    completionPhrase:
      optionalString(fields[columns.CompletionPhrase]) || defaults.completionPhrase,
    welcomeMessage:
      optionalString(fields[columns.WelcomeMessage]) || defaults.welcomeMessage,
    locale: requiredString(fields[columns.Locale], "Locale", itemId),
    voiceName: requiredString(fields[columns.VoiceName], "VoiceName", itemId),
    avatarCharacter: requiredString(
      fields[columns.AvatarCharacter],
      "AvatarCharacter",
      itemId,
    ),
    avatarStyle: requiredString(fields[columns.AvatarStyle], "AvatarStyle", itemId),
  };
}

export async function loadAgentCatalog(
  config: AppConfig,
  identity: MicrosoftIdentity,
): Promise<AgentDefinition[]> {
  if (!config.sharePointCatalog) {
    return [config.defaultAgent];
  }

  const token = await identity.acquireToken(graphScopes);
  const { hostname, sitePath, listName } = config.sharePointCatalog;
  const normalizedPath = `/${sitePath.replace(/^\/+|\/+$/g, "")}`;
  const site = await graphGet<GraphSite>(
    `${graphRoot}/sites/${encodeURIComponent(hostname)}:${normalizedPath}`,
    token,
  );
  if (!site.id) {
    throw new Error("Microsoft Graph did not return the configured SharePoint site.");
  }

  const escapedListName = listName.replace(/'/g, "''");
  const lists = await graphGet<GraphListResponse>(
    `${graphRoot}/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName&$filter=${encodeURIComponent(
      `displayName eq '${escapedListName}'`,
    )}`,
    token,
  );
  const list = lists.value?.find((candidate) => candidate.displayName === listName);
  if (!list?.id) {
    throw new Error(`SharePoint list "${listName}" was not found.`);
  }

  const columnResult = await graphGet<GraphColumnsResponse>(
    `${graphRoot}/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(
      list.id,
    )}/columns?$select=name,displayName`,
    token,
  );
  const columns = resolveColumns(columnResult.value || []);
  const fieldSelection = Object.values(columns).join(",");
  const agents: AgentDefinition[] = [];
  let nextUrl: string | undefined =
    `${graphRoot}/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(
      list.id,
    )}/items?$expand=fields($select=${fieldSelection})&$top=200`;

  for (let page = 0; nextUrl && page < 5; page += 1) {
    const result: GraphItemsResponse = await graphGet<GraphItemsResponse>(nextUrl, token);
    for (const item of result.value || []) {
      const agent = mapAgent(item, columns, config.defaultAgent);
      if (agent) {
        agents.push(agent);
      }
    }
    nextUrl = result["@odata.nextLink"];
  }

  if (agents.length === 0) {
    throw new Error(`SharePoint list "${listName}" has no enabled agents.`);
  }

  return agents.sort((left, right) => left.displayName.localeCompare(right.displayName));
}
