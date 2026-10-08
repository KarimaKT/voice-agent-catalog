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

import type {
  AgentAction,
  AgentAttachment,
  AgentCitation,
  AgentDefinition,
  AppConfig,
} from "./types";

function createConnectionSettings(agent: AgentDefinition): ConnectionSettings {
  return new ConnectionSettings({
    environmentId: agent.environmentId,
    schemaName: agent.schemaName,
    cloud: "Prod",
    copilotAgentType: "Published",
  });
}

export interface AgentTurn {
  messages: AgentMessage[];
  conversationId: string;
  emailDraft?: {
    to: string;
    subject: string;
    body: string;
  };
}

export interface AgentMessage {
  text: string;
  speak?: string;
  attachments?: AgentAttachment[];
  citations?: AgentCitation[];
  suggestedActions?: AgentAction[];
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
        const wrapped = new Error(message);
        (wrapped as Error & { cause?: unknown }).cause = interactiveError;
        throw wrapped;
      }
    }
  }

  getBestEffortUserEmail(): string | undefined {
    return this.teamsContext.loginHint || this.account?.username || undefined;
  }

  async ensureUserIdentity(agent: AgentDefinition): Promise<string | undefined> {
    const settings = createConnectionSettings(agent);
    await this.acquireToken([ScopeHelper.getScopeFromSettings(settings)]);
    return this.getBestEffortUserEmail();
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
    const messages: AgentMessage[] = [];

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
    const messages: AgentMessage[] = [];

    for await (const reply of this.client.sendActivityStreaming(activity, this.conversationId)) {
      this.captureActivity(reply, messages);
    }

    return { messages, conversationId: this.conversationId };
  }

  private captureActivity(activity: Activity, messages: AgentMessage[]): void {
    if (activity.conversation?.id) {
      this.conversationId = activity.conversation.id;
    }
    if (activity.type === ActivityTypes.Message) {
      const message = normalizeActivity(activity);
      if (
        message.text ||
        message.attachments?.length ||
        message.citations?.length ||
        message.suggestedActions?.length
      ) {
        messages.push(message);
      }
    }
  }
}

function safeHttpUrl(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  try {
    const url = new URL(value);
    return url.protocol === "http:" || url.protocol === "https:" ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function textValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function adaptiveCardText(value: unknown): string {
  if (Array.isArray(value)) {
    return value.map(adaptiveCardText).filter(Boolean).join(" ");
  }
  if (!value || typeof value !== "object") {
    return "";
  }
  const item = value as Record<string, unknown>;
  const ownText =
    textValue(item.text) ||
    (item.type === "FactSet" && Array.isArray(item.facts)
      ? item.facts
          .map((fact) => {
            if (!fact || typeof fact !== "object") {
              return "";
            }
            const entry = fact as Record<string, unknown>;
            return [textValue(entry.title), textValue(entry.value)].filter(Boolean).join(" ");
          })
          .join(" ")
      : "");
  return [
    ownText,
    adaptiveCardText(item.items),
    adaptiveCardText(item.columns),
  ].filter(Boolean).join(" ");
}

function normalizeActivity(activity: Activity): AgentMessage {
  const attachments: AgentAttachment[] = [];
  const attachmentSpeech: string[] = [];
  for (const item of activity.attachments || []) {
    const contentType = item.contentType.toLocaleLowerCase();
    const url = safeHttpUrl(item.contentUrl);
    if (contentType === "application/vnd.microsoft.card.adaptive" && item.content) {
      const card = item.content as Record<string, unknown>;
      attachments.push({
        kind: "adaptiveCard",
        name: textValue(item.name),
        body: Array.isArray(card.body)
          ? card.body.filter((entry): entry is Record<string, unknown> =>
              Boolean(entry && typeof entry === "object"),
            )
          : [],
        actions: Array.isArray(card.actions)
          ? card.actions.filter((entry): entry is Record<string, unknown> =>
              Boolean(entry && typeof entry === "object"),
            )
          : [],
      });
      const cardSpeech = adaptiveCardText(card.body);
      if (cardSpeech) {
        attachmentSpeech.push(cardSpeech);
      }
    } else if (url && contentType.startsWith("image/")) {
      attachments.push({
        kind: "image",
        url,
        name: textValue(item.name),
        alt: textValue(item.name) || textValue(activity.summary) || "Agent-generated image",
      });
    } else if (url) {
      attachments.push({
        kind: "file",
        url,
        name: textValue(item.name) || "Open attachment",
        contentType: item.contentType,
      });
    }
  }

  const citations: AgentCitation[] = [];
  for (const entity of activity.entities || []) {
    const candidate = entity as unknown as Record<string, unknown>;
    const rawCitations = candidate.citation;
    if (!Array.isArray(rawCitations)) {
      continue;
    }
    for (const raw of rawCitations) {
      if (!raw || typeof raw !== "object") {
        continue;
      }
      const appearance = (raw as Record<string, unknown>).appearance;
      if (!appearance || typeof appearance !== "object") {
        continue;
      }
      const detail = appearance as Record<string, unknown>;
      const name = textValue(detail.name);
      if (name) {
        citations.push({
          name,
          abstract: textValue(detail.abstract),
          url: safeHttpUrl(detail.url),
        });
      }
    }
  }

  const suggestedActions: AgentAction[] = (activity.suggestedActions?.actions || [])
    .flatMap<AgentAction>((action) => {
      const title = textValue(action.title) || textValue(action.displayText);
      if (!title) {
        return [];
      }
      const type = String(action.type || "").toLocaleLowerCase();
      const value = textValue(action.value);
      if (type === "openurl") {
        const url = safeHttpUrl(value);
        return url ? [{ title, url }] : [];
      }
      return [{ title, value: value || title }];
    });

  return {
    text: activity.text?.trim() || "",
    speak:
      activity.speak?.trim() ||
      (!activity.text?.trim()
        ? activity.summary?.trim() || attachmentSpeech.join(" ")
        : undefined) ||
      undefined,
    attachments: attachments.length ? attachments : undefined,
    citations: citations.length ? citations : undefined,
    suggestedActions: suggestedActions.length ? suggestedActions : undefined,
  };
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
  private pendingOrder?: Omit<DemoOrder, "orderNumber" | "status">;

  constructor(
    private readonly agent: AgentDefinition,
    private readonly emailRecipient?: string,
  ) {}

  async connect(): Promise<AgentTurn> {
    this.answers = [];
    if (this.agent.demoKind === "orders") {
      return {
        conversationId: this.conversationId,
        messages: [{ text: this.agent.welcomeMessage }],
      };
    }
    return {
      conversationId: this.conversationId,
      messages: [{
        text: `Hi, I'm ${this.agent.displayName}, your manager handoff assistant. This is demo mode, so no Microsoft sign-in is required. ${demoQuestions[0]}`,
      }],
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
        messages: [{ text: this.runOrderTool(answer) }],
      };
    }
    this.answers.push(answer);
    const nextQuestion = demoQuestions[this.answers.length];
    if (nextQuestion) {
      return {
        conversationId: this.conversationId,
        messages: [{ text: `Thank you. ${nextQuestion}` }],
      };
    }

    return {
      conversationId: this.conversationId,
      messages: [{
        text: [
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
      }],
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
    const normalized = query.trim().toLocaleLowerCase();
    if (this.pendingOrder) {
      if (/^(?:yes|y|confirm|confirmed|place it|create it|go ahead|do it)\b/i.test(normalized)) {
        const order: DemoOrder = {
          ...this.pendingOrder,
          orderNumber: `ORD-${1000 + this.orders.length + 1}`,
          status: "Pending",
        };
        this.orders.push(order);
        this.pendingOrder = undefined;
        return [
          this.formatOrders([order], "Order created"),
          "This demo tool stores the order for the current session. You can continue managing orders.",
        ].join("\n\n");
      }
      if (/^(?:no|n|cancel|stop|never mind|nevermind)\b/i.test(normalized)) {
        this.pendingOrder = undefined;
        return "I canceled that order request. What would you like to do next?";
      }
      return "Please confirm the pending order with “yes”, or cancel it with “no”.";
    }

    const wantsOrder = /\b(?:place|create|add|submit|new)\b.*\border\b|\border\b.*\b(?:for|of)\b/i.test(query);
    if (wantsOrder) {
      const parsed = this.parseOrderRequest(query);
      if (!parsed) {
        return "Tell me the quantity, product, and customer. For example: “Create an order for 4 Surface Laptop 7 for Contoso Retail.”";
      }
      this.pendingOrder = parsed;
      return `Please confirm: ${parsed.quantity} x ${parsed.product} for ${parsed.customer}. Should I create this order?`;
    }

    const orderNumber = query.match(/\bORD[\s-]?(\d+)\b/i)?.[1];
    if (orderNumber) {
      const canonicalOrderNumber = `ORD-${orderNumber}`;
      const order = this.orders.find(
        (candidate) => candidate.orderNumber === canonicalOrderNumber,
      );
      return order
        ? this.formatOrders([order], `Order ${canonicalOrderNumber}`)
        : `I couldn't find ${canonicalOrderNumber}. Try another order number.`;
    }

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

  private parseOrderRequest(
    query: string,
  ): Omit<DemoOrder, "orderNumber" | "status"> | undefined {
    const patterns = [
      /(?:place|create|add|submit)(?:\s+(?:a|an|new))?\s+order\s+(?:for|of)\s+(\d+)\s+(.+?)\s+for\s+(.+?)(?:[.!?]|$)/i,
      /(?:place|create|add|submit)(?:\s+(?:a|an|new))?\s+order\s+for\s+(.+?),\s*(?:quantity\s*)?(\d+),?\s*(?:customer\s*)?(.+?)(?:[.!?]|$)/i,
      /(\d+)\s+(.+?)\s+for\s+(.+?)\s+(?:please|order|ordered)?(?:[.!?]|$)/i,
    ];
    for (const [index, pattern] of patterns.entries()) {
      const match = query.match(pattern);
      if (!match) {
        continue;
      }
      const quantity = Number(index === 1 ? match[2] : match[1]);
      const product = (index === 1 ? match[1] : match[2]).trim();
      const customer = match[3].trim();
      if (Number.isInteger(quantity) && quantity > 0 && product && customer) {
        return { quantity, product, customer };
      }
    }
    return undefined;
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
