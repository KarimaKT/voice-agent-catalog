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
  demoKind?: "interview" | "orders";
}

export interface AppConfig {
  tenantId: string;
  clientId: string;
  defaultAgent: AgentDefinition;
  catalogEnabled: boolean;
  demoMode: boolean;
  demoAgents?: AgentDefinition[];
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
