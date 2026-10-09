export interface AgentDefinition {
  id: string;
  displayName: string;
  description?: string;
  environmentId: string;
  schemaName: string;
  completionPhrase: string;
  endsConversation: boolean;
  welcomeMessage: string;
  locale: string;
  voiceName: string;
  avatarCharacter: string;
  avatarStyle: string;
  audience?: "Everyone" | "Restricted";
  harness?: string;
  demoKind?: "interview" | "orders";
}

export interface AppConfig {
  tenantId: string;
  clientId: string;
  apiScope: string;
  defaultAgent: AgentDefinition;
  catalogEnabled: boolean;
  demoMode: boolean;
  demoAgents?: AgentDefinition[];
  costNotices: {
    voice: string;
    avatar: string;
  };
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
  kind?: "message" | "submit";
  value?: unknown;
  url?: string;
  unsupported?: string;
}

export interface AgentCitation {
  name: string;
  abstract?: string;
  url?: string;
}

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
      content: Record<string, unknown>;
    }
  | {
      kind: "unsupported";
      name?: string;
      reason: string;
    };

export type TurnState =
  | "connecting"
  | "ready"
  | "listening"
  | "thinking"
  | "speaking"
  | "complete"
  | "error";
