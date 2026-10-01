export interface AgentDefinition {
  id: string;
  displayName: string;
  description?: string;
  environmentId: string;
  schemaName: string;
  completionPhrase: string;
  welcomeMessage: string;
  locale: string;
  voiceName: string;
  avatarCharacter: string;
  avatarStyle: string;
}

export interface SharePointCatalogConfig {
  hostname: string;
  sitePath: string;
  listName: string;
}

export interface AppConfig {
  tenantId: string;
  clientId: string;
  defaultAgent: AgentDefinition;
  sharePointCatalog?: SharePointCatalogConfig;
}

export interface SpeechCredentials {
  token: string;
  region: string;
  relay?: {
    url: string;
    username: string;
    credential: string;
  };
}

export interface ChatMessage {
  id: string;
  role: "agent" | "user" | "system";
  text: string;
}

export type TurnState =
  | "connecting"
  | "ready"
  | "listening"
  | "thinking"
  | "speaking"
  | "complete"
  | "error";
