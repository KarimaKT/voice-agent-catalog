import {
  createNestablePublicClientApplication,
  createStandardPublicClientApplication,
  type AccountInfo,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { Activity, ActivityTypes } from "@microsoft/agents-activity";
import {
  ConnectionSettings,
  CopilotStudioClient,
  ScopeHelper,
  type StartRequest,
} from "@microsoft/agents-copilotstudio-client";

import type { AgentDefinition, AppConfig } from "./types";

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
  emailDraft?: {
    to: string;
    subject: string;
    body: string;
  };
}

export interface AgentClient {
  connect(): Promise<AgentTurn>;
  send(text: string): Promise<AgentTurn>;
}

export interface TeamsIdentityContext {
  isTeamsHosted: boolean;
  supportsNestedAuth: boolean;
  homeAccountId?: string;
  loginHint?: string;
  tenantId?: string;
}

export class MicrosoftIdentity {
  private readonly msalPromise: Promise<IPublicClientApplication>;
  private account?: AccountInfo;

  constructor(config: AppConfig, private readonly teamsContext: TeamsIdentityContext) {
    const auth = {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      ...(teamsContext.isTeamsHosted
        ? {}
        : { redirectUri: `${window.location.origin}/auth/callback` }),
    };
    const msalConfig = {
      auth: {
        ...auth,
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    } as const;

    this.msalPromise = teamsContext.supportsNestedAuth
      ? createNestablePublicClientApplication(msalConfig)
      : createStandardPublicClientApplication(msalConfig);
  }

  async acquireToken(scopes: string[], interactive = true): Promise<string> {
    if (this.teamsContext.isTeamsHosted && !this.teamsContext.supportsNestedAuth) {
      throw new Error(
        "This Teams client does not support silent app authentication. Update Teams and try again.",
      );
    }

    const msal = await this.msalPromise;
    this.account ||= msal.getActiveAccount() ?? undefined;
    this.account ||=
      msal.getAccount({
        homeAccountId: this.teamsContext.homeAccountId,
        loginHint: this.teamsContext.loginHint,
        tenantId: this.teamsContext.tenantId,
      }) ?? undefined;
    if (this.account) {
      msal.setActiveAccount(this.account);
    }

    const request = {
      account: this.account,
      scopes,
    };

    try {
      const result = await msal.acquireTokenSilent(request);
      this.account = result.account;
      return result.accessToken;
    } catch (error) {
      if (!interactive) {
        throw error;
      }
      try {
        const result = await msal.acquireTokenPopup({
          scopes,
          loginHint: this.teamsContext.loginHint,
        });
        this.account = result.account;
        return result.accessToken;
      } catch (interactiveError) {
        const message = this.teamsContext.isTeamsHosted
          ? "Teams could not authorize Copilot Studio. Grant the requested permission and try again; if the tenant requires admin approval, ask an administrator or use the anonymous Direct Line option."
          : "Microsoft sign-in was cancelled or could not authorize Copilot Studio.";
        throw new Error(message, { cause: interactiveError });
      }
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

const demoQuestions = [
  "What is your name and role?",
  "What are your top priorities right now?",
  "What blockers or support do you need from your manager?",
  "What decisions, updates, or next steps should your manager know about?",
];

interface DemoOrder {
  orderNumber: string;
  customer: string;
  product: string;
  quantity: number;
  status: "Pending" | "Processing" | "Shipped";
}

const initialDemoOrders: DemoOrder[] = [
  {
    orderNumber: "ORD-1001",
    customer: "Contoso Retail",
    product: "Surface Laptop 7",
    quantity: 5,
    status: "Processing",
  },
  {
    orderNumber: "ORD-1002",
    customer: "Fabrikam Foods",
    product: "Surface Pro 11",
    quantity: 12,
    status: "Pending",
  },
  {
    orderNumber: "ORD-1003",
    customer: "Northwind Traders",
    product: "Teams Rooms Pro",
    quantity: 3,
    status: "Shipped",
  },
  {
    orderNumber: "ORD-1004",
    customer: "Adventure Works",
    product: "Copilot for Microsoft 365",
    quantity: 25,
    status: "Pending",
  },
];

export class DemoAgentClient implements AgentClient {
  private readonly conversationId = crypto.randomUUID();
  private answers: string[] = [];
  private orders = initialDemoOrders.map((order) => ({ ...order }));

  constructor(
    private readonly agent: AgentDefinition,
    private readonly emailRecipient?: string,
  ) {}

  async connect(): Promise<AgentTurn> {
    this.answers = [];
    if (this.agent.demoKind === "orders") {
      return {
        conversationId: this.conversationId,
        messages: [this.agent.welcomeMessage],
      };
    }
    return {
      conversationId: this.conversationId,
      messages: [
        `Hi, I'm ${this.agent.displayName}, your manager handoff assistant. This is demo mode, so no Microsoft sign-in is required. ${demoQuestions[0]}`,
      ],
    };
  }

  async send(text: string): Promise<AgentTurn> {
    const answer = text.trim();
    if (!answer) {
      throw new Error("Enter an answer before continuing.");
    }
    if (this.agent.demoKind === "orders") {
      return {
        conversationId: this.conversationId,
        messages: [this.runOrderTool(answer)],
      };
    }
    this.answers.push(answer);
    const nextQuestion = demoQuestions[this.answers.length];
    if (nextQuestion) {
      return {
        conversationId: this.conversationId,
        messages: [`Thank you. ${nextQuestion}`],
      };
    }

    return {
      conversationId: this.conversationId,
      messages: [
        [
          this.agent.completionPhrase,
          "",
          "Manager handoff summary",
          `Name and role: ${this.answers[0]}`,
          `Top priorities: ${this.answers[1]}`,
          `Blockers and support: ${this.answers[2]}`,
          `Decisions and next steps: ${this.answers[3]}`,
          "",
          this.emailRecipient
            ? "Select Open self-addressed email, review the Outlook draft, and select Send."
            : "Email is not configured in demo mode.",
        ].join("\n"),
      ],
      emailDraft: this.emailRecipient
        ? {
            to: this.emailRecipient,
            subject: "Manager handoff summary",
            body: [
              "Manager handoff summary",
              "",
              `Name and role: ${this.answers[0]}`,
              `Top priorities: ${this.answers[1]}`,
              `Blockers and support: ${this.answers[2]}`,
              `Decisions and next steps: ${this.answers[3]}`,
            ].join("\n"),
          }
        : undefined,
    };
  }

  private runOrderTool(query: string): string {
    const orderNumber = query.match(/ORD-\d+/i)?.[0].toUpperCase();
    if (orderNumber) {
      const order = this.orders.find((candidate) => candidate.orderNumber === orderNumber);
      return order
        ? this.formatOrders([order], `Order ${orderNumber}`)
        : `I couldn't find ${orderNumber}. Try another order number.`;
    }

    const placeMatch = query.match(
      /(?:place|create)\s+(?:an?\s+)?order\s+(?:for\s+)?(\d+)\s+(.+?)\s+for\s+(.+)/i,
    );
    if (placeMatch) {
      const order: DemoOrder = {
        orderNumber: `ORD-${1000 + this.orders.length + 1}`,
        quantity: Number(placeMatch[1]),
        product: placeMatch[2].trim(),
        customer: placeMatch[3].trim(),
        status: "Pending",
      };
      this.orders.push(order);
      return [
        this.agent.completionPhrase,
        this.formatOrders([order], "New order"),
        "This demo tool stores the order for the current session.",
      ].join("\n\n");
    }

    const normalized = query.toLocaleLowerCase();
    if (normalized.includes("pending") || normalized.includes("open")) {
      return this.formatOrders(
        this.orders.filter((order) => order.status === "Pending"),
        "Pending orders",
      );
    }
    if (normalized.includes("all") || normalized.includes("show") || normalized.includes("list")) {
      return this.formatOrders(this.orders, "All orders");
    }

    return [
      "I can use the order tool in three ways:",
      '- "Show pending orders"',
      '- "Find order ORD-1002"',
      '- "Place an order for 4 Surface Laptop 7 for Contoso Retail"',
    ].join("\n");
  }

  private formatOrders(orders: DemoOrder[], heading: string): string {
    if (orders.length === 0) {
      return `${heading}: none found.`;
    }
    return [
      heading,
      ...orders.map(
        (order) =>
          `${order.orderNumber} | ${order.customer} | ${order.quantity} x ${order.product} | ${order.status}`,
      ),
    ].join("\n");
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
