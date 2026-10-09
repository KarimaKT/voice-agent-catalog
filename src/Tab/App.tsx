import React from "react";
import * as teamsJs from "@microsoft/teams-js";

import { CatalogAccessError, loadAgentCatalog } from "./catalog";
import {
  CopilotAgentClient,
  DemoAgentClient,
  MicrosoftIdentity,
  toSpokenText,
  type AgentClient,
  type AgentMessage,
  type TeamsIdentityContext,
} from "./copilot";
import { SpeechController } from "./speech";
import { completesConversation, getHarnessSupport } from "./agent-behavior";
import { AdaptiveCard } from "./AdaptiveCard";
import { safeHttpUrl } from "./adaptive-card-policy";
import type {
  AgentAction,
  AgentAttachment,
  AgentDefinition,
  AvatarIdleTimeoutSeconds,
  AppConfig,
  ChatMessage,
  TurnState,
} from "./types";
import "./App.css";

const inputLanguages = [
  ["en-US", "English (United States)"],
  ["en-GB", "English (United Kingdom)"],
  ["es-ES", "Spanish (Spain)"],
  ["fr-FR", "French (France)"],
  ["de-DE", "German (Germany)"],
  ["it-IT", "Italian (Italy)"],
  ["pt-BR", "Portuguese (Brazil)"],
  ["hi-IN", "Hindi (India)"],
  ["ja-JP", "Japanese (Japan)"],
  ["ko-KR", "Korean (Korea)"],
  ["zh-CN", "Chinese (Mandarin, Simplified)"],
] as const;

function createMessage(
  role: ChatMessage["role"],
  content: string | AgentMessage,
): ChatMessage {
  const message = typeof content === "string" ? { text: content } : content;
  return {
    id: `${Date.now()}-${crypto.randomUUID()}`,
    role,
    ...message,
  };
}

function RichMessageContent({
  message,
  onAction,
  actionsDisabled,
}: {
  message: ChatMessage;
  onAction: (action: AgentAction) => void;
  actionsDisabled: boolean;
}) {
  return (
    <>
      {message.text && <p>{message.text}</p>}
      {message.attachments?.map((attachment, index) => (
        <AttachmentContent
          attachment={attachment}
          key={`${attachment.kind}-${index}`}
          onAction={onAction}
          actionsDisabled={actionsDisabled}
        />
      ))}
      {message.citations?.length ? (
        <div className="citations" aria-label="Sources">
          <strong>Sources</strong>
          <ol>
            {message.citations.map((citation, index) => (
              <li key={`${citation.name}-${index}`}>
                {safeHttpUrl(citation.url) ? (
                  <a href={safeHttpUrl(citation.url)} target="_blank" rel="noopener noreferrer">
                    {citation.name}
                  </a>
                ) : citation.name}
                {citation.abstract && <span>{citation.abstract}</span>}
              </li>
            ))}
          </ol>
        </div>
      ) : null}
      {message.suggestedActions?.length ? (
        <div className="message-actions">
          {message.suggestedActions.map((action, index) => (
            action.unsupported ? (
              <p className="card-fallback" role="note" key={`${action.title}-${index}`}>
                {action.title}: {action.unsupported}
              </p>
            ) : (
              <button
                className="action-button"
                key={`${action.title}-${index}`}
                disabled={!action.url && actionsDisabled}
                onClick={() => onAction(action)}
              >
                {action.title}
              </button>
            )
          ))}
        </div>
      ) : null}
    </>
  );
}

function AttachmentContent({
  attachment,
  onAction,
  actionsDisabled,
}: {
  attachment: AgentAttachment;
  onAction: (action: AgentAction) => void;
  actionsDisabled: boolean;
}) {
  if (attachment.kind === "unsupported") {
    return <p className="card-fallback" role="note">{attachment.name}: {attachment.reason}</p>;
  }
  if (attachment.kind === "image") {
    const url = safeHttpUrl(attachment.url);
    if (!url) return <p role="note">This image link was blocked.</p>;
    return (
      <figure className="message-image">
        <img src={url} alt={attachment.alt || attachment.name || "Agent-generated image"} loading="lazy" />
        {attachment.name && <figcaption>{attachment.name}</figcaption>}
      </figure>
    );
  }
  if (attachment.kind === "file") {
    const url = safeHttpUrl(attachment.url);
    if (!url) return <p role="note">This attachment link was blocked.</p>;
    return (
      <a className="file-link" href={url} target="_blank" rel="noopener noreferrer">
        {attachment.name}
      </a>
    );
  }
  return (
    <AdaptiveCard
      content={attachment.content}
      name={attachment.name}
      disabled={actionsDisabled}
      onAction={onAction}
    />
  );
}

async function loadConfig(): Promise<AppConfig> {
  const response = await fetch("/api/config");
  if (!response.ok) {
    throw new Error("The app is not configured. Check the environment settings.");
  }
  return (await response.json()) as AppConfig;
}

export default function App() {
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [draft, setDraft] = React.useState("");
  const [interim, setInterim] = React.useState("");
  const [state, setState] = React.useState<TurnState>("connecting");
  const [status, setStatus] = React.useState("Preparing the app...");
  const [avatarEnabled, setAvatarEnabled] = React.useState(false);
  const [readRepliesAloud, setReadRepliesAloud] = React.useState(false);
  const [avatarIdleSeconds, setAvatarIdleSeconds] = React.useState<AvatarIdleTimeoutSeconds>(34);
  const [costNotices, setCostNotices] = React.useState<AppConfig["costNotices"]>();
  const [avatarVisible, setAvatarVisible] = React.useState(false);
  const [voiceConversationActive, setVoiceConversationActive] = React.useState(false);
  const [hostName, setHostName] = React.useState("browser");
  const [agents, setAgents] = React.useState<AgentDefinition[]>([]);
  const [catalogLoaded, setCatalogLoaded] = React.useState(false);
  const [selectedAgentId, setSelectedAgentId] = React.useState("");
  const [emailDraftUrl, setEmailDraftUrl] = React.useState("");
  const [inputLanguage, setInputLanguage] = React.useState("en-US");
  const [canManageCatalog, setCanManageCatalog] = React.useState(false);
  const [configurationUrl, setConfigurationUrl] = React.useState("");
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const configRef = React.useRef<AppConfig | undefined>(undefined);
  const demoEmailRecipientRef = React.useRef<string | undefined>(undefined);
  const identityRef = React.useRef<MicrosoftIdentity | undefined>(undefined);
  const agentClientRef = React.useRef<AgentClient | undefined>(undefined);
  const speechRef = React.useRef<SpeechController | undefined>(undefined);
  const transcriptRef = React.useRef<HTMLDivElement>(null);
  const voiceConversationActiveRef = React.useRef(false);
  const avatarEnabledRef = React.useRef(false);
  const readRepliesAloudRef = React.useRef(false);
  const autoSubmitTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const turnInFlightRef = React.useRef(false);
  const catalogLoadInFlightRef = React.useRef(false);
  const sessionEpochRef = React.useRef(0);

  React.useEffect(() => {
    let disposed = false;

    async function initialize(): Promise<void> {
      let identityContext: TeamsIdentityContext = {
        isTeamsHosted: false,
        supportsNestedAuth: false,
      };
      try {
        await teamsJs.app.initialize();
        const context = await teamsJs.app.getContext();
        let supportsNestedAuth = false;
        try {
          supportsNestedAuth = await teamsJs.nestedAppAuth.isNAAChannelRecommended();
        } catch {
          supportsNestedAuth = false;
        }
        identityContext = {
          isTeamsHosted: true,
          supportsNestedAuth,
          loginHint: context.user?.loginHint,
          tenantId: context.user?.tenant?.id,
        };
        if (context.app.host.name) {
          setHostName(context.app.host.name);
        }
      } catch {
        setHostName("browser");
      }

      try {
        const config = await loadConfig();
        if (!videoRef.current || disposed) {
          return;
        }
        configRef.current = config;
        setCostNotices(config.costNotices);
        demoEmailRecipientRef.current = identityContext.loginHint;
        const identity = new MicrosoftIdentity(config, identityContext);
        identityRef.current = identity;
        speechRef.current = new SpeechController(
          config.defaultAgent,
          videoRef.current,
          () => identity.acquireApiToken(),
          (message) => {
            setAvatarVisible(false);
            setMessages((current) => [...current, createMessage("system", message)]);
          },
        );
        setInputLanguage(config.defaultAgent.locale);
        const initialAgents = config.demoMode ? config.demoAgents || [] : [];
        setAgents(initialAgents);
        setCatalogLoaded(initialAgents.length > 0);
        setSelectedAgentId(
          (initialAgents.find((agent) => agent.id === config.defaultAgent.id &&
            getHarnessSupport(agent.harness).supported) ||
            initialAgents.find((agent) => getHarnessSupport(agent.harness).supported))?.id || "",
        );
        setState("ready");
        setStatus(initialAgents.length ? "Choose an agent to connect" : "Sign in to load your authorized catalog");
        if (await identity.completeRedirect() && !disposed) {
          void loadCatalog();
        }
      } catch (error) {
        setState("error");
        setStatus(error instanceof Error ? error.message : "Initialization failed.");
      }
    }

    void initialize();
    const releaseOnExit = () => {
      sessionEpochRef.current++;
      setVoiceConversation(false);
      void speechRef.current?.close().catch(reportCleanupError);
      setAvatarVisible(false);
      setState((current) =>
        current === "complete" || current === "error" || current === "thinking" ? current : "ready");
      setStatus("Audio/video released while away. Restart voice when ready.");
    };
    const releaseWhenHidden = () => {
      if (document.hidden) releaseOnExit();
    };
    window.addEventListener("pagehide", releaseOnExit);
    document.addEventListener("visibilitychange", releaseWhenHidden);
    return () => {
      disposed = true;
      sessionEpochRef.current++;
      voiceConversationActiveRef.current = false;
      window.removeEventListener("pagehide", releaseOnExit);
      document.removeEventListener("visibilitychange", releaseWhenHidden);
      if (autoSubmitTimerRef.current) {
        clearTimeout(autoSubmitTimerRef.current);
      }
      void speechRef.current?.close().catch(reportCleanupError);
    };
  }, []);

  React.useEffect(() => {
    transcriptRef.current?.scrollTo({
      top: transcriptRef.current.scrollHeight,
      behavior: "smooth",
    });
  }, [messages, interim]);

  function reportCleanupError(error: unknown): void {
    const message = error instanceof Error ? error.message : "Speech cleanup failed.";
    setMessages((current) => [...current, createMessage("system", `Speech cleanup: ${message}`)]);
  }

  async function loadCatalog(): Promise<void> {
    const config = configRef.current;
    const identity = identityRef.current;
    if (!config || !identity || catalogLoadInFlightRef.current) {
      return;
    }
    catalogLoadInFlightRef.current = true;
    setState("connecting");
    setStatus("Signing in and loading your authorized catalog...");
    setCanManageCatalog(false);
    setConfigurationUrl("");
    try {
      const catalog = await loadAgentCatalog(config, await identity.acquireApiToken());
      setAgents(catalog.agents);
      setCanManageCatalog(catalog.canManageCatalog);
      setConfigurationUrl(catalog.configurationUrl || "");
      const selected = catalog.agents.find((agent) => getHarnessSupport(agent.harness).supported);
      setSelectedAgentId(selected?.id || "");
      if (selected) setInputLanguage(selected.locale);
      setCatalogLoaded(true);
      setState("ready");
      setStatus(selected
        ? "Catalog loaded. Choose an agent, then connect or start voice."
        : "Catalog loaded. No agents use the supported Standard harness transport.");
    } catch (error) {
      setCatalogLoaded(false);
      handleError(error instanceof CatalogAccessError
        ? error
        : new Error(error instanceof Error
          ? `The shared agent catalog is unavailable: ${error.message}`
          : "The shared agent catalog is unavailable."), "Could not load your authorized catalog.");
    } finally {
      catalogLoadInFlightRef.current = false;
    }
  }

  async function connect(): Promise<void> {
    if (!catalogLoaded) {
      await loadCatalog();
      return;
    }
    const config = configRef.current;
    const identity = identityRef.current;
    if (!config || !identity) return;
    const epoch = ++sessionEpochRef.current;
    const selected = agents.find((agent) => agent.id === selectedAgentId);
    setState("connecting");
    setStatus(`Connecting to ${selected?.displayName || "the selected agent"}...`);
    try {
      if (!selected || !getHarnessSupport(selected.harness).supported) {
        throw new Error("Choose an authorized agent with the supported Standard harness transport.");
      }
      setSelectedAgentId(selected.id);
      await speechRef.current?.setConversationActive(false);
      await speechRef.current?.setAgent(selected);
      speechRef.current?.setInputLanguage(inputLanguage);
      setAvatarVisible(false);
      const client = config.demoMode
        ? new DemoAgentClient(selected, demoEmailRecipientRef.current)
        : new CopilotAgentClient(identity, selected);
      const turn = await client.connect();
      if (epoch !== sessionEpochRef.current) return;
      agentClientRef.current = client;
      const completed = turn.completed === true ||
        turn.messages.some((message) => completesConversation(selected, message.text));
      const welcome =
        turn.messages.length > 0
          ? turn.messages
          : [{ text: completed ? "Conversation complete." : selected.welcomeMessage }];
      const narrate = readRepliesAloudRef.current || avatarEnabledRef.current || voiceConversationActiveRef.current;
      setMessages(welcome.map((message) => createMessage("agent", message)));
      if (completed) setVoiceConversation(false);
      await speechRef.current?.setConversationActive(!completed && !config.demoMode);
      setState(completed ? "complete" : "ready");
      setStatus(completed ? "Conversation complete" : `${selected.displayName} is ready`);
      await speakMessages(
        welcome
          .map((message) => message.speak || toSpokenText(message.text))
          .filter(Boolean),
        completed,
        epoch,
        narrate,
      );
    } catch (error) {
      if (epoch !== sessionEpochRef.current) return;
      agentClientRef.current = undefined;
      handleError(error, `Could not connect to ${selected?.displayName || "the selected agent"}.`);
    }
  }

  async function sendMessage(inputOverride?: string | AgentAction): Promise<void> {
    const input = inputOverride ?? draft;
    const text = typeof input === "string" ? input.trim() : input.title;
    if (!text || !agentClientRef.current || turnInFlightRef.current ||
        state === "complete" || state === "connecting" || state === "speaking") {
      return;
    }
    turnInFlightRef.current = true;
    speechRef.current?.setAgentWorking(true);
    const epoch = sessionEpochRef.current;
    const client = agentClientRef.current;
    if (autoSubmitTimerRef.current) {
      clearTimeout(autoSubmitTimerRef.current);
      autoSubmitTimerRef.current = undefined;
    }
    try {
      await speechRef.current?.stopListening().catch(() => undefined);
      setInterim("");
      setDraft("");
      setMessages((current) => [...current, createMessage("user", text)]);
      setState("thinking");
      setStatus(`${selectedAgent?.displayName || "The agent"} is thinking...`);
      const turn = await client.send(typeof input === "string" ? text : input);
      if (client !== agentClientRef.current) return;
      if (turn.messages.length === 0 && !turn.completed) {
        throw new Error(`${selectedAgent?.displayName || "The agent"} returned no message.`);
      }

      const nextMessages: ChatMessage[] = [];
      const spoken: string[] = [];
      let completed = turn.completed === true;
      const completionAgent =
        agents.find((agent) => agent.id === selectedAgentId) ||
        configRef.current?.defaultAgent;

      for (const reply of turn.messages) {
        completed ||= Boolean(
          completionAgent && completesConversation(completionAgent, reply.text),
        );
        nextMessages.push(createMessage("agent", reply));
        const speech = reply.speak || toSpokenText(reply.text);
        if (speech) {
          spoken.push(speech);
        }
        if (completed && nextMessages.length === 0) {
          nextMessages.push(createMessage("system", "Conversation complete."));
        }
      }

      setMessages((current) => [...current, ...nextMessages]);
      setEmailDraftUrl(
        turn.emailDraft
          ? `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(turn.emailDraft.to)}&subject=${encodeURIComponent(turn.emailDraft.subject)}&body=${encodeURIComponent(turn.emailDraft.body)}`
          : "",
      );
      const narrate = readRepliesAloudRef.current || avatarEnabledRef.current || voiceConversationActiveRef.current;
      if (completed) {
        setVoiceConversation(false);
        await speechRef.current?.setConversationActive(false);
        setAvatarVisible(false);
        setState("complete");
        setStatus("Conversation complete");
        await speakMessages(spoken, true, epoch, narrate);
      } else {
        if (epoch === sessionEpochRef.current && !document.hidden) {
          await speechRef.current?.setConversationActive(!configRef.current?.demoMode);
        }
        setState("ready");
        setStatus("Your turn");
        await speakMessages(spoken, false, epoch, narrate);
      }
    } catch (error) {
      handleError(error, `${selectedAgent?.displayName || "The agent"} could not complete that turn.`);
    } finally {
      turnInFlightRef.current = false;
      speechRef.current?.setAgentWorking(false);
    }
  }

  async function startListening(autoSubmit = false): Promise<void> {
    if (!speechRef.current) {
      return;
    }
    try {
      setInterim("");
      setState("listening");
      setStatus(
        autoSubmit
          ? "Listening... pause when you finish"
          : "Listening... select I'm done when finished",
      );
      await speechRef.current.startListening(
        (text) => {
          setInterim(text);
          if (autoSubmitTimerRef.current) {
            clearTimeout(autoSubmitTimerRef.current);
          }
        },
        (text) => {
          setInterim(text);
          setDraft(text);
          if (autoSubmit && voiceConversationActiveRef.current) {
            if (autoSubmitTimerRef.current) {
              clearTimeout(autoSubmitTimerRef.current);
            }
            autoSubmitTimerRef.current = setTimeout(() => {
              void submitVoiceAnswer(text);
            }, 1400);
          }
        },
        (message) => {
          setVoiceConversation(false);
          handleError(new Error(message), "Speech recognition stopped.");
        },
      );
    } catch (error) {
      handleError(error, "Microphone access failed.");
    }
  }

  async function submitVoiceAnswer(text: string): Promise<void> {
    if (!voiceConversationActiveRef.current || !text.trim()) {
      return;
    }
    if (autoSubmitTimerRef.current) {
      clearTimeout(autoSubmitTimerRef.current);
      autoSubmitTimerRef.current = undefined;
    }
    await speechRef.current?.stopListening();
    setInterim("");
    await sendMessage(text);
  }

  function setVoiceConversation(active: boolean): void {
    voiceConversationActiveRef.current = active;
    setVoiceConversationActive(active);
    if (!active && autoSubmitTimerRef.current) {
      clearTimeout(autoSubmitTimerRef.current);
      autoSubmitTimerRef.current = undefined;
    }
  }

  async function startVoiceConversation(): Promise<void> {
    if (!catalogLoaded) {
      await loadCatalog();
      return;
    }
    setVoiceConversation(true);
    if (!agentClientRef.current) {
      await connect();
      return;
    }
    await speechRef.current?.setConversationActive(!configRef.current?.demoMode);
    await startListening(true);
  }

  async function stopVoiceConversation(): Promise<void> {
    sessionEpochRef.current++;
    setVoiceConversation(false);
    try {
      await speechRef.current?.setConversationActive(false);
    } catch (error) {
      reportCleanupError(error);
    }
    setAvatarVisible(false);
    if (state === "thinking") {
      setStatus("Voice stopped; waiting for the in-flight agent reply.");
    } else if (state !== "complete" && state !== "error") {
      setState("ready");
      setStatus("Voice conversation stopped");
    }
  }

  async function speakMessages(
    texts: string[],
    preserveCompletion = false,
    epoch = sessionEpochRef.current,
    narrate = readRepliesAloudRef.current || avatarEnabledRef.current || voiceConversationActiveRef.current,
  ): Promise<void> {
    if (!speechRef.current || epoch !== sessionEpochRef.current || document.hidden) {
      return;
    }
    const useAvatar = avatarEnabledRef.current && !preserveCompletion && !configRef.current?.demoMode;
    if (!narrate) return;
    if (texts.length === 0) {
      if (voiceConversationActiveRef.current && !preserveCompletion) {
        await startListening(true);
      }
      return;
    }
    setState(preserveCompletion ? "complete" : "speaking");
    setStatus(`${selectedAgent?.displayName || "The agent"} is speaking...`);
    try {
      for (const text of texts) {
        if (epoch !== sessionEpochRef.current || document.hidden) return;
        await speechRef.current.speak(text, useAvatar && avatarEnabledRef.current);
      }
    } catch (error) {
      if (epoch !== sessionEpochRef.current || document.hidden) return;
      if (useAvatar) {
        avatarEnabledRef.current = false;
        setAvatarEnabled(false);
        setMessages((current) => [
          ...current,
          createMessage("system", "Avatar video is unavailable; continuing with audio-only mode."),
        ]);
        for (const text of texts) {
          if (epoch !== sessionEpochRef.current) return;
          await speechRef.current.speak(text, false);
        }
        if (epoch !== sessionEpochRef.current || document.hidden) return;
        setAvatarVisible(false);
      } else {
        handleError(error, "Speech playback failed.");
        return;
      }
    }
    setState(preserveCompletion ? "complete" : "ready");
    setStatus(preserveCompletion ? "Conversation complete" : "Your turn");
    if (voiceConversationActiveRef.current && !preserveCompletion) {
      await startListening(true);
    }
  }

  function handleError(error: unknown, fallback: string): void {
    sessionEpochRef.current++;
    setVoiceConversation(false);
    agentClientRef.current = undefined;
    void speechRef.current?.setConversationActive(false).catch(reportCleanupError);
    setAvatarVisible(false);
    const message = error instanceof Error ? error.message : fallback;
    setState("error");
    setStatus(message || fallback);
    setMessages((current) => [...current, createMessage("system", message || fallback)]);
  }

  async function selectAgent(agentId: string): Promise<void> {
    if (agentId === selectedAgentId) {
      return;
    }
    sessionEpochRef.current++;
    await speechRef.current?.setConversationActive(false);
    setAvatarVisible(false);
    setVoiceConversation(false);
    agentClientRef.current = undefined;
    setMessages([]);
    setDraft("");
    setInterim("");
    setEmailDraftUrl("");
    setSelectedAgentId(agentId);
    const selected = agents.find((agent) => agent.id === agentId);
    if (selected) {
      await speechRef.current?.setAgent(selected);
      setInputLanguage(selected.locale);
      speechRef.current?.setInputLanguage(selected.locale);
    }
    setState("ready");
    setStatus(selected ? `Ready to connect to ${selected.displayName}` : "Ready to connect");
  }

  const connected = Boolean(agentClientRef.current);
  const busy = state === "connecting" || state === "thinking" || state === "speaking";
  const selectedAgent = agents.find((agent) => agent.id === selectedAgentId) || agents[0];
  const selectedHarness = getHarnessSupport(selectedAgent?.harness);

  function handleAgentAction(action: AgentAction): void {
    if (action.unsupported) return;
    if (action.url) {
      const url = safeHttpUrl(action.url);
      if (url) window.open(url, "_blank", "noopener,noreferrer");
      return;
    }
    if (action.kind === "submit" || typeof action.value === "string") {
      void sendMessage(action);
    }
  }

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">
            {selectedAgent ? `Meet ${selectedAgent.displayName}` : "Voice-enabled Copilot Studio"}
          </p>
          <h1>Voice Agent Catalog</h1>
          <p className="subtitle">
            {selectedAgent?.description || "A short guided interview with voice and avatar."}
          </p>
        </div>
        <div className={`status status-${state}`} aria-live="polite">
          <span className="status-dot" />
          {status}
        </div>
      </header>

      <section className="workspace">
        <div className="conversation-panel">
          <div className="panel-heading">
            <div>
              <h2>Conversation</h2>
              <p>Running in {hostName}</p>
            </div>
            <label className="agent-picker">
              <span>Choose an agent</span>
              <select
                value={selectedAgentId}
                onChange={(event) => void selectAgent(event.target.value)}
                disabled={!catalogLoaded || busy || state === "listening"}
              >
                {!selectedAgentId && <option value="">
                  {catalogLoaded ? "No supported agents available" : "Load catalog to choose an agent"}
                </option>}
                {agents.map((agent) => (
                  <option
                    value={agent.id}
                    key={agent.id}
                    disabled={!getHarnessSupport(agent.harness).supported}
                  >
                    {agent.displayName}
                    {!getHarnessSupport(agent.harness).supported
                      ? ` — ${getHarnessSupport(agent.harness).label} transport unavailable`
                      : ""}
                  </option>
                ))}
              </select>
              {selectedAgent && !selectedHarness.supported && (
                <small className="transport-message">
                  {selectedHarness.label} harness transport is not supported. Choose a Standard agent.
                </small>
              )}
            </label>
            <label className="agent-picker">
              <span>Spoken input</span>
              <select
                value={inputLanguage}
                onChange={(event) => {
                  setInputLanguage(event.target.value);
                  speechRef.current?.setInputLanguage(event.target.value);
                }}
                disabled={busy || state === "listening"}
              >
                {!inputLanguages.some(([locale]) => locale === inputLanguage) && (
                  <option value={inputLanguage}>{inputLanguage}</option>
                )}
                {inputLanguages.map(([locale, label]) => (
                  <option value={locale} key={locale}>{label}</option>
                ))}
              </select>
            </label>
            {canManageCatalog && configurationUrl && (
              <a className="catalog-admin-link" href={configurationUrl} target="_blank" rel="noreferrer">
                Manage catalog
              </a>
            )}
          </div>

          <p className="voice-interface-note">
            AI responses are requested in the configured output language with concise,
            voice-friendly wording. Agent-authored cards and controls may affect the
            experience. Attachments are displayed in chat, not read aloud.
          </p>
          <section className="speech-cost-notice" aria-label="Voice and avatar costs">
            <strong>Choose paid audio/video consciously</strong>
            <p>{costNotices?.voice || "Voice pricing is loading; check the rates before starting."}</p>
            <p>{costNotices?.avatar || "Avatar pricing is loading; check the rates before enabling video."}</p>
            <small>
              Estimates, not a live bill. Hosting, Copilot Studio and licensing are additional.
              Charges go to the deployment owner's services. Text-only mode uses no Speech.
              Avatar is reused during an active Copilot Studio conversation and disconnects
              after the selected idle period. The idle timer pauses while the agent works
              or speaks; connected waiting time remains billable.
            </small>
          </section>
          <div className="transcript" ref={transcriptRef} aria-live="polite">
            {!connected && (
              <div className="welcome-card">
                <div className="pat-mark">{selectedAgent?.displayName.charAt(0) || "A"}</div>
                <h3>Ready when you are</h3>
                <p>
                  {hostName === "browser"
                    ? "Connect with your work account for browser diagnostics. "
                    : configRef.current?.demoMode
                      ? "Demo mode starts without login or consent. "
                      : "Teams uses your work identity automatically; first use may request Copilot Studio permission once. "}
                  You can type or use your microphone, and you can turn avatar video off
                  at any time.
                </p>
                <button
                  className="primary-button"
                  onClick={() => void (catalogLoaded ? connect() : loadCatalog())}
                  disabled={busy || (catalogLoaded && !selectedHarness.supported)}
                >
                  {!catalogLoaded
                    ? "Sign in / Load catalog"
                    : `${configRef.current?.demoMode ? "Start demo with" : "Connect to"} ${selectedAgent?.displayName || "agent"}`}
                </button>
              </div>
            )}

            {messages.map((message) => (
              <article className={`message message-${message.role}`} key={message.id}>
                <span className="message-author">
                  {message.role === "agent"
                    ? selectedAgent?.displayName || "Agent"
                    : message.role === "user"
                      ? "You"
                      : "Status"}
                </span>
                <RichMessageContent
                  message={message}
                  onAction={handleAgentAction}
                  actionsDisabled={busy || state === "complete" || !agentClientRef.current}
                />
              </article>
            ))}

            {interim && state === "listening" && (
              <article className="message message-interim">
                <span className="message-author">Listening</span>
                <p>{interim}</p>
              </article>
            )}
          </div>

          <div className="composer">
            <label htmlFor="answer">Your answer</label>
            <textarea
              id="answer"
              value={draft}
              onChange={(event) => {
                setDraft(event.target.value);
                speechRef.current?.noteUserActivity();
              }}
              placeholder={
                connected
                  ? "Type a message or use the microphone..."
                  : `Connect to ${selectedAgent?.displayName || "an agent"} first`
              }
              disabled={!catalogLoaded || !selectedHarness.supported || busy || state === "complete"}
              rows={3}
              onKeyDown={(event) => {
                if (event.key === "Enter" && !event.shiftKey) {
                  event.preventDefault();
                  void sendMessage();
                }
              }}
            />
            <div className="composer-actions">
              {voiceConversationActive ? (
                <button
                  className={`secondary-button${state === "listening" ? " listening" : ""}`}
                  onClick={() => void stopVoiceConversation()}
                >
                  Stop conversation
                </button>
              ) : (
                <button
                  className="secondary-button"
                  onClick={() => void startVoiceConversation()}
                  disabled={!catalogLoaded || !selectedHarness.supported || busy || state === "complete"}
                >
                  Start voice conversation
                </button>
              )}
              <button
                className="primary-button"
                onClick={() => void sendMessage()}
                disabled={!draft.trim() || !connected || busy || state === "complete"}
              >
                Send
              </button>
              {emailDraftUrl && (
                <a
                  className="primary-button"
                  href={emailDraftUrl}
                  target="_blank"
                  rel="noreferrer"
                >
                  Open self-addressed email
                </a>
              )}
            </div>
          </div>
        </div>

        <aside className="avatar-panel">
          <div className="avatar-stage">
            <video
              ref={videoRef}
              className={avatarVisible ? "avatar-visible" : undefined}
              autoPlay
              playsInline
              onPlaying={() => setAvatarVisible(true)}
              onEmptied={() => setAvatarVisible(false)}
              aria-label={`${selectedAgent?.displayName || "Agent"} speaking avatar`}
            />
            <div className="avatar-placeholder">
              <div className="avatar-orb">
                <span>{selectedAgent?.displayName.charAt(0) || "A"}</span>
              </div>
              <p>
                {avatarEnabled
                  ? `${selectedAgent?.displayName || "The agent"}'s avatar appears here`
                  : "Audio-only mode"}
              </p>
            </div>
          </div>

          <div className="avatar-controls">
            <label className="toggle">
              <input
                type="checkbox"
                checked={readRepliesAloud}
                disabled={!costNotices}
                onChange={(event) => {
                  readRepliesAloudRef.current = event.target.checked;
                  setReadRepliesAloud(event.target.checked);
                }}
              />
              <span>Read typed replies aloud (paid voice)</span>
            </label>
            <label className="toggle">
              <input
                type="checkbox"
                checked={avatarEnabled}
                disabled={!costNotices || configRef.current?.demoMode}
                onChange={(event) => {
                  avatarEnabledRef.current = event.target.checked;
                  setAvatarEnabled(event.target.checked);
                  if (!event.target.checked) {
                    setAvatarVisible(false);
                    void speechRef.current?.stopAvatar().catch(reportCleanupError);
                  }
                }}
              />
              <span>Avatar video</span>
            </label>
            <button
              className="text-button"
              onClick={() => void speechRef.current?.stopSpeaking().catch(reportCleanupError)}
              disabled={state !== "speaking"}
            >
              Stop speaking
            </button>
          </div>

          <label className="agent-picker">
            <span>Avatar idle timeout</span>
            <select
              value={avatarIdleSeconds}
              disabled={!costNotices}
              onChange={(event) => {
                const seconds = Number(event.target.value);
                if (seconds !== 15 && seconds !== 34 && seconds !== 45) {
                  handleError(new Error("Choose 15, 34 or 45 seconds."), "Invalid avatar timeout.");
                  return;
                }
                speechRef.current?.setAvatarIdleTimeout(seconds);
                setAvatarIdleSeconds(seconds);
              }}
            >
              <option value={15}>15 seconds</option>
              <option value={34}>34 seconds</option>
              <option value={45}>45 seconds</option>
            </select>
            <small>
              Only counts idle time waiting for you, not agent processing or spoken replies.
              Stops video only; the next reply can reconnect it.
            </small>
          </label>

          <div className="privacy-note">
            <h3>You stay in control</h3>
            <p>
              Use text or speech, turn avatar video on or off, and stop playback at any
              time. Each agent follows its own configured purpose and workflow.
            </p>
            {selectedAgent && (
              <p className="agent-identity">
                Voice: {selectedAgent.voiceName} · Avatar: {selectedAgent.avatarCharacter} (
                {selectedAgent.avatarStyle}) · Output: {selectedAgent.locale} · Input: {inputLanguage}
              </p>
            )}
          </div>

          {state === "complete" && (
            <div className="completion-card" role="status">
              <strong>Complete</strong>
              <span>{selectedAgent?.completionPhrase}</span>
            </div>
          )}
        </aside>
      </section>
    </main>
  );
}
