import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";
import { Activity, ActivityTypes } from "@microsoft/agents-activity";
import {
  ConnectionSettings,
  CopilotStudioClient,
  ScopeHelper,
  type StartRequest,
} from "@microsoft/agents-copilotstudio-client";

import type { AgentDefinition, AppConfig } from "./types";

const graphScopes = ["https://graph.microsoft.com/Sites.Read.All"];

function createConnectionSettings(agent: AgentDefinition): ConnectionSettings {
  return new ConnectionSettings({
    environmentId: agent.environmentId,
    schemaName: agent.schemaName,
    cloud: "Prod",
    copilotAgentType: "Published",
  });
}

export interface AgentTurn {
  messages: string[];
  conversationId: string;
}

export class MicrosoftIdentity {
  private readonly msal: PublicClientApplication;
  private account?: AccountInfo;

  constructor(config: AppConfig, private readonly loginHint?: string) {
    this.msal = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: `${window.location.origin}/auth/callback`,
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    });
  }

  async signIn(): Promise<void> {
    await this.msal.initialize();
    const accounts = this.msal.getAllAccounts();
    this.account = accounts[0];

    if (!this.account) {
      const login = await this.msal.loginPopup({
        scopes: ["openid", "profile", ...graphScopes],
        loginHint: this.loginHint,
      });
      this.account = login.account ?? undefined;
    }

    if (!this.account) {
      throw new Error("Microsoft sign-in did not return an account.");
    }
  }

  async acquireToken(scopes: string[]): Promise<string> {
    await this.signIn();
    if (!this.account) {
      throw new Error("A Microsoft account is required.");
    }

    const request = {
      account: this.account,
      scopes,
    };

    try {
      return (await this.msal.acquireTokenSilent(request)).accessToken;
    } catch {
      return (await this.msal.acquireTokenPopup(request)).accessToken;
    }
  }
}

export class CopilotAgentClient {
  private readonly settings: ConnectionSettings;
  private client?: CopilotStudioClient;
  private conversationId = "";

  constructor(
    private readonly identity: MicrosoftIdentity,
    private readonly agent: AgentDefinition,
  ) {
    this.settings = createConnectionSettings(agent);
  }

  async connect(): Promise<AgentTurn> {
    const token = await this.identity.acquireToken([
      ScopeHelper.getScopeFromSettings(this.settings),
    ]);
    this.client = new CopilotStudioClient(this.settings, token);
    const startRequest: StartRequest = {
      emitStartConversationEvent: false,
      locale: this.agent.locale,
    };
    const messages: string[] = [];

    for await (const activity of this.client.startConversationStreaming(startRequest)) {
      this.captureActivity(activity, messages);
    }

    if (!this.conversationId) {
      throw new Error(`${this.agent.displayName} did not return a conversation ID.`);
    }

    return { messages, conversationId: this.conversationId };
  }

  async send(text: string): Promise<AgentTurn> {
    if (!this.client || !this.conversationId) {
      throw new Error(`Connect to ${this.agent.displayName} before sending a message.`);
    }

    const activity = new Activity(ActivityTypes.Message);
    activity.text = text;
    activity.conversation = { id: this.conversationId };
    const messages: string[] = [];

    for await (const reply of this.client.sendActivityStreaming(activity, this.conversationId)) {
      this.captureActivity(reply, messages);
    }

    return { messages, conversationId: this.conversationId };
  }

  private captureActivity(activity: Activity, messages: string[]): void {
    if (activity.conversation?.id) {
      this.conversationId = activity.conversation.id;
    }
    if (activity.type === ActivityTypes.Message && activity.text?.trim()) {
      messages.push(activity.text.trim());
    }
  }
}

export function toSpokenText(value: string): string {
  return value
    .replace(/```[\s\S]*?```/g, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/https?:\/\/\S+/g, "")
    .replace(/[*_#>`]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
