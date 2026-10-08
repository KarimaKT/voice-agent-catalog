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
  audience?: "Everyone" | "Restricted";
  allowedUsers?: CatalogUser[];
  harness?: string;
  demoKind?: "interview" | "orders";
}

export interface CatalogUser {
  email: string;
  displayName?: string;
}

export interface AppConfig {
  tenantId: string;
  clientId: string;
  defaultAgent: AgentDefinition;
  catalogEnabled: boolean;
  demoMode: boolean;
  demoAgents?: AgentDefinition[];
}

export interface CatalogResult {
  agents: AgentDefinition[];
  canManageCatalog: boolean;
  configurationUrl?: string;
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
  attachments?: AgentAttachment[];
  citations?: AgentCitation[];
  suggestedActions?: AgentAction[];
}

export interface AgentAction {
  title: string;
  value?: string;
  url?: string;
}

export interface AgentCitation {
  name: string;
  abstract?: string;
  url?: string;
}

export type AdaptiveCardElement = Record<string, unknown>;

export type AgentAttachment =
  | {
      kind: "image";
      url: string;
      name?: string;
      alt?: string;
    }
  | {
      kind: "file";
      url: string;
      name: string;
      contentType?: string;
    }
  | {
      kind: "adaptiveCard";
      name?: string;
      body: AdaptiveCardElement[];
      actions: AdaptiveCardElement[];
    };

export type TurnState =
  | "connecting"
  | "ready"
  | "listening"
  | "thinking"
  | "speaking"
  | "complete"
  | "error";
