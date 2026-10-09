import fs from "fs";
import https from "https";
import path from "path";

import { useAzureMonitor } from "@azure/monitor-opentelemetry";
import { App, ExpressAdapter } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { type NextFunction, type Request, type Response } from "express";
import { authenticateApiRequest } from "./auth";
import { validateAndAuthorizeCatalog } from "./catalog-policy";

interface RelayTokenResponse {
  Urls?: string[];
  Username?: string;
  Password?: string;
}

interface RateLimitEntry {
  count: number;
  resetAt: number;
}

interface CatalogFlowResponse {
  agents?: unknown[];
  canManageCatalog?: unknown;
  configurationUrl?: unknown;
}

if (process.env.APPLICATIONINSIGHTS_CONNECTION_STRING) {
  useAzureMonitor({
    samplingRatio: 0,
    tracesPerSecond: 0,
    enableLiveMetrics: false,
    browserSdkLoaderOptions: { enabled: false },
    instrumentationOptions: {
      http: { enabled: false },
      azureSdk: { enabled: false },
      console: { enabled: false },
      bunyan: { enabled: false },
      winston: { enabled: false },
      mongoDb: { enabled: false },
      mySql: { enabled: false },
      postgreSql: { enabled: false },
      redis: { enabled: false },
      redis4: { enabled: false },
    },
  });
}

const requiredEnvironment = [
  "TENANT_ID",
  "AAD_APP_CLIENT_ID",
  "COPILOT_ENVIRONMENT_ID",
  "COPILOT_SCHEMA_NAME",
  "SPEECH_KEY",
  "SPEECH_REGION",
  "SPEECH_ENDPOINT",
] as const;

const defaultCompletionPhrase =
  "Thank you. Your manager handoff summary has been emailed.";
const defaultWelcomeMessage =
  "Hi, I'm Pat, an AI assistant to your manager. I'll ask four short questions and email your manager a handoff summary. What is your name and role?";

function safeHttpUrl(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  try {
    const url = new URL(value);
    return url.protocol === "https:" ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function getEnvironment(name: (typeof requiredEnvironment)[number]): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : undefined,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : undefined,
};

const secureServer =
  sslOptions.cert && sslOptions.key ? https.createServer(sslOptions) : undefined;
const adapter = secureServer
  ? new ExpressAdapter(secureServer)
  : new ExpressAdapter();
adapter.use((_request: Request, response: Response, next: NextFunction) => {
  response.removeHeader("X-Powered-By");
  response.setHeader("Content-Security-Policy", "default-src 'self'; connect-src 'self' https://login.microsoftonline.com https://res.cdn.office.net https://api.powerplatform.com https://*.api.powerplatform.com https://*.environment.api.powerplatform.com https://*.cognitiveservices.azure.com https://*.speech.microsoft.com wss://*.speech.microsoft.com; frame-src 'self' https://login.microsoftonline.com; frame-ancestors 'self' https://teams.microsoft.com https://*.teams.microsoft.com https://*.cloud.microsoft; img-src 'self' https: data: blob:; media-src 'self' blob:; script-src 'self'; worker-src 'self' blob: data:; style-src 'self' 'unsafe-inline'");
  response.setHeader("Referrer-Policy", "no-referrer");
  response.setHeader("X-Content-Type-Options", "nosniff");
  next();
});
const logger = new ConsoleLogger("voice-agent-catalog", { level: "info" });
const app = new App({
  logger,
  httpServerAdapter: adapter,
  dangerouslyAllowUnauthenticatedRequests: true,
});
app.event("error", ({ error }) => {
  const message = error instanceof Error ? error.message : "Application server error";
  logger.error(message);
  console.error(message);
  process.exitCode = 1;
});

const speechRateLimits = new Map<string, RateLimitEntry>();

function enforceSpeechRateLimit(userId: string, response: Response): boolean {
  const now = Date.now();
  for (const [id, entry] of speechRateLimits) {
    if (entry.resetAt <= now) {
      speechRateLimits.delete(id);
    }
  }
  const key = userId;
  const existing = speechRateLimits.get(key);
  const entry = !existing || existing.resetAt <= now
    ? { count: 0, resetAt: now + 60_000 }
    : existing;

  entry.count += 1;
  speechRateLimits.set(key, entry);
  response.setHeader("RateLimit-Limit", "30");
  response.setHeader("RateLimit-Remaining", String(Math.max(0, 30 - entry.count)));
  response.setHeader("RateLimit-Reset", String(Math.ceil((entry.resetAt - now) / 1000)));

  if (entry.count <= 30) {
    return true;
  }

  response.status(429).json({ error: "Too many Speech requests. Try again shortly." });
  return false;
}

adapter.get("/api/health", (_request: Request, response: Response) => {
  response.setHeader("Cache-Control", "no-store");
  response.json({ status: "ok" });
});

adapter.get("/", (_request: Request, response: Response) => {
  response.setHeader("Cache-Control", "no-store");
  response.redirect(302, "/tabs/home");
});

adapter.get("/api/config", (_request: Request, response: Response) => {
  try {
    const demoMode = process.env.DEMO_MODE?.trim().toLowerCase() === "true";
    const defaultAgent = {
      id: "default",
      displayName: process.env.DEFAULT_AGENT_DISPLAY_NAME?.trim() || "Pat",
      description:
        process.env.DEFAULT_AGENT_DESCRIPTION?.trim() ||
        "A four-question manager handoff interview in under three minutes.",
      environmentId: getEnvironment("COPILOT_ENVIRONMENT_ID"),
      schemaName: getEnvironment("COPILOT_SCHEMA_NAME"),
      completionPhrase:
        process.env.DEFAULT_AGENT_COMPLETION_PHRASE?.trim() ||
        (demoMode
          ? "Thank you. Your manager handoff summary is ready."
          : defaultCompletionPhrase),
      endsConversation: true,
      welcomeMessage:
        process.env.DEFAULT_AGENT_WELCOME_MESSAGE?.trim() ||
        defaultWelcomeMessage,
      locale: process.env.SPEECH_LOCALE?.trim() || "en-US",
      voiceName:
        process.env.SPEECH_VOICE_NAME?.trim() || "en-US-AvaMultilingualNeural",
      avatarCharacter: process.env.AVATAR_CHARACTER?.trim() || "lisa",
      avatarStyle: process.env.AVATAR_STYLE?.trim() || "casual-sitting",
      audience: "Everyone",
      harness: "standard",
      demoKind: "interview",
    };
    const catalogFlowUrl = process.env.CATALOG_FLOW_URL?.trim();
    const catalogEnabled =
      Boolean(catalogFlowUrl) && catalogFlowUrl?.toLowerCase() !== "disabled";
    const clientId = getEnvironment("AAD_APP_CLIENT_ID");
    response.setHeader("Cache-Control", "no-store");
    response.json({
      tenantId: getEnvironment("TENANT_ID"),
      clientId,
      apiScope: `api://${clientId}/access_as_user`,
      defaultAgent,
      catalogEnabled,
      demoMode,
      costNotices: {
        voice: process.env.VOICE_COST_NOTICE?.trim() ||
          "Voice estimate (USD): input $1/hour of audio sent (~$0.0167/min); output $15/million characters ($0.015/1,000 characters).",
        avatar: process.env.AVATAR_COST_NOTICE?.trim() ||
          "Avatar estimate (USD): $0.50/min while connected, including silence, plus voice synthesis. Three connected minutes cost $1.50 for video.",
      },
      demoAgents: demoMode
        ? [
            defaultAgent,
            {
              ...defaultAgent,
              id: "orders",
              displayName: "Morgan",
              description:
                "An order management assistant that finds orders and creates new order requests.",
              completionPhrase: "Your order request has been placed.",
              endsConversation: false,
              welcomeMessage:
                "Hi, I'm Morgan. Ask me to show pending orders, find an order number, or place a new order.",
              voiceName: "en-US-AndrewMultilingualNeural",
              avatarCharacter: "harry",
              avatarStyle: "casual",
              demoKind: "orders",
            },
          ]
        : undefined,
    });
  } catch (error) {
    logger.error(error instanceof Error ? error.message : "Unable to load app configuration");
    response.status(500).json({ error: "The application is not configured." });
  }
});

adapter.get("/api/catalog", async (request: Request, response: Response) => {
  let user;
  try {
    const tenantId = getEnvironment("TENANT_ID");
    const audience =
      process.env.API_AUDIENCE?.trim() || getEnvironment("AAD_APP_CLIENT_ID");
    user = await authenticateApiRequest(request, tenantId, audience);
  } catch (error) {
    logger.warn(error instanceof Error ? error.message : "Catalog authentication failed");
    response.status(401).json({ error: "A valid app access token is required." });
    return;
  }

  const catalogFlowUrl = process.env.CATALOG_FLOW_URL?.trim();
  if (!catalogFlowUrl || catalogFlowUrl.toLowerCase() === "disabled") {
    response.status(404).json({ error: "The shared agent catalog is not configured." });
    return;
  }
  if (!user.email) {
    response.status(403).json({ error: "The app token must contain a verified user email or UPN for catalog access." });
    return;
  }

  try {
    const flowResponse = await fetch(catalogFlowUrl, {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        operation: "listEnabledAgents",
        userEmail: user.email,
        userObjectId: user.oid,
      }),
      signal: AbortSignal.timeout(20_000),
    });
    if (!flowResponse.ok) {
      throw new Error(`Catalog flow failed with status ${flowResponse.status}`);
    }
    const payload = (await flowResponse.json()) as CatalogFlowResponse;
    if (!Array.isArray(payload.agents)) {
      throw new Error("Catalog flow returned an invalid response");
    }
    const agents = validateAndAuthorizeCatalog(payload.agents, user);

    const canManageCatalog = payload.canManageCatalog === true;
    response.setHeader("Cache-Control", "no-store");
    response.json({
      agents,
      canManageCatalog,
      configurationUrl: canManageCatalog
        ? safeHttpUrl(payload.configurationUrl)
        : undefined,
    });
  } catch (error) {
    logger.error(error instanceof Error ? error.message : "Unable to load agent catalog");
    response.status(502).json({ error: "The shared agent catalog is temporarily unavailable." });
  }
});

adapter.get("/auth/callback", (_request: Request, response: Response) => {
  response.setHeader("Cache-Control", "no-store");
  response.sendFile(path.join(__dirname, "client", "auth-callback.html"));
});

adapter.post("/api/speech/token", async (request: Request, response: Response) => {
  let user;
  try {
    const tenantId = getEnvironment("TENANT_ID");
    const audience =
      process.env.API_AUDIENCE?.trim() || getEnvironment("AAD_APP_CLIENT_ID");
    user = await authenticateApiRequest(request, tenantId, audience);
  } catch (error) {
    logger.warn(error instanceof Error ? error.message : "Speech authentication failed");
    response.status(401).json({ error: "A valid app access token is required." });
    return;
  }

  if (!enforceSpeechRateLimit(user.oid, response)) {
    return;
  }

  try {
    const speechKey = getEnvironment("SPEECH_KEY");
    const speechRegion = getEnvironment("SPEECH_REGION");
    const speechEndpoint = new URL(getEnvironment("SPEECH_ENDPOINT"));
    const tokenUrl = `https://${speechRegion}.api.cognitive.microsoft.com/sts/v1.0/issueToken`;
    const relayUrl = new URL("/tts/cognitiveservices/avatar/relay/token/v1", speechEndpoint);

    const [tokenResult, relayResult] = await Promise.all([
      fetch(tokenUrl, {
        method: "POST",
        headers: { "Ocp-Apim-Subscription-Key": speechKey },
        signal: AbortSignal.timeout(15_000),
      }),
      fetch(relayUrl, {
        headers: { "Ocp-Apim-Subscription-Key": speechKey },
        signal: AbortSignal.timeout(15_000),
      }),
    ]);

    if (!tokenResult.ok) {
      throw new Error(`Speech token request failed with status ${tokenResult.status}`);
    }
    const token = await tokenResult.text();
    let relayCredentials:
      | { url: string; username: string; credential: string }
      | undefined;

    if (relayResult.ok) {
      const relay = (await relayResult.json()) as RelayTokenResponse;
      const turnUrl = relay.Urls?.find((url) => url.startsWith("turn:")) ?? relay.Urls?.[0];
      if (turnUrl && relay.Username && relay.Password) {
        relayCredentials = {
          url: turnUrl,
          username: relay.Username,
          credential: relay.Password,
        };
      }
    } else {
      logger.warn(`Avatar relay request failed with status ${relayResult.status}`);
    }

    response.setHeader("Cache-Control", "no-store");
    response.json({
      token,
      region: speechRegion,
      relay: relayCredentials,
    });
  } catch (error) {
    logger.error(error instanceof Error ? error.message : "Unable to issue Speech credentials");
    response.status(502).json({ error: "Speech is temporarily unavailable." });
  }
});

adapter.get("/privacy", (_request: Request, response: Response) => {
  response.type("html").send("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>Voice Agent Catalog Privacy</title></head><body><main><h1>Voice Agent Catalog privacy</h1><p>This proof of concept sends conversation messages to the selected Copilot Studio agent and uses Azure Speech for optional voice and avatar features. It does not intentionally persist conversation transcripts in this application.</p></main></body></html>");
});

adapter.get("/terms", (_request: Request, response: Response) => {
  response.type("html").send("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>Voice Agent Catalog Terms</title></head><body><main><h1>Voice Agent Catalog terms</h1><p>This application is a proof of concept for authorized demonstration users. Do not submit sensitive personal, medical, legal, compensation, or confidential production information.</p></main></body></html>");
});

app.tab("home", path.join(__dirname, "./client"));

async function start(): Promise<void> {
  requiredEnvironment.forEach(getEnvironment);
  const port = Number(process.env.PORT || 3978);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("PORT must be a valid TCP port number.");
  }
  logger.info(`Starting Voice Agent Catalog on TCP port ${port}`);
  await app.start(port);
  logger.info(`Voice Agent Catalog is listening on TCP port ${port}`);
}

void start().catch((error: unknown) => {
  logger.error(error instanceof Error ? error.message : "Application startup failed");
  process.exitCode = 1;
});
