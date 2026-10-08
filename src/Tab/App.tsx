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
import type {
  AdaptiveCardElement,
  AgentAction,
  AgentAttachment,
  AgentDefinition,
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

function safeHttpUrl(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  try {
    const url = new URL(value);
    return url.protocol === "https:" || url.protocol === "http:" ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function cardText(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function AdaptiveCardContent({
  elements,
  onAction,
}: {
  elements: AdaptiveCardElement[];
  onAction: (action: AgentAction) => void;
}) {
  return (
    <>
      {elements.map((element, index) => {
        const type = cardText(element.type);
        const key = `${type || "element"}-${index}`;
        if (type === "TextBlock") {
          return <p className="adaptive-text" key={key}>{cardText(element.text)}</p>;
        }
        if (type === "Image") {
          const url = safeHttpUrl(element.url);
          return url ? (
            <img
              className="adaptive-image"
              key={key}
              src={url}
              alt={cardText(element.altText) || "Adaptive Card image"}
              loading="lazy"
            />
          ) : null;
        }
        if (type === "FactSet" && Array.isArray(element.facts)) {
          return (
            <dl className="adaptive-facts" key={key}>
              {element.facts.map((fact, factIndex) => {
                const item = fact && typeof fact === "object"
                  ? fact as Record<string, unknown>
                  : {};
                return (
                  <React.Fragment key={`${key}-${factIndex}`}>
                    <dt>{cardText(item.title)}</dt>
                    <dd>{cardText(item.value)}</dd>
                  </React.Fragment>
                );
              })}
            </dl>
          );
        }
        if (
          (type === "Container" || type === "Column") &&
          Array.isArray(element.items)
        ) {
          return (
            <div className="adaptive-container" key={key}>
              <AdaptiveCardContent elements={element.items as AdaptiveCardElement[]} onAction={onAction} />
            </div>
          );
        }
        if (type === "ColumnSet" && Array.isArray(element.columns)) {
          return (
            <div className="adaptive-columns" key={key}>
              {element.columns.map((column, columnIndex) => {
                const item = column && typeof column === "object"
                  ? column as AdaptiveCardElement
                  : {};
                return (
                  <AdaptiveCardContent
                    key={`${key}-${columnIndex}`}
                    elements={[item]}
                    onAction={onAction}
                  />
                );
              })}
            </div>
          );
        }
        return null;
      })}
    </>
  );
}

function RichMessageContent({
  message,
  onAction,
}: {
  message: ChatMessage;
  onAction: (action: AgentAction) => void;
}) {
  return (
    <>
      {message.text && <p>{message.text}</p>}
      {message.attachments?.map((attachment, index) => (
        <AttachmentContent
          attachment={attachment}
          key={`${attachment.kind}-${index}`}
          onAction={onAction}
        />
      ))}
      {message.citations?.length ? (
        <div className="citations" aria-label="Sources">
          <strong>Sources</strong>
          <ol>
            {message.citations.map((citation, index) => (
              <li key={`${citation.name}-${index}`}>
                {citation.url ? (
                  <a href={citation.url} target="_blank" rel="noreferrer">
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
            <button
              className="action-button"
              key={`${action.title}-${index}`}
              onClick={() => onAction(action)}
            >
              {action.title}
            </button>
          ))}
        </div>
      ) : null}
    </>
  );
}

function AttachmentContent({
  attachment,
  onAction,
}: {
  attachment: AgentAttachment;
  onAction: (action: AgentAction) => void;
}) {
  if (attachment.kind === "image") {
    return (
      <figure className="message-image">
        <img src={attachment.url} alt={attachment.alt || attachment.name || "Agent-generated image"} loading="lazy" />
        {attachment.name && <figcaption>{attachment.name}</figcaption>}
      </figure>
    );
  }
  if (attachment.kind === "file") {
    return (
      <a className="file-link" href={attachment.url} target="_blank" rel="noreferrer">
        {attachment.name}
      </a>
    );
  }
  return (
    <section className="adaptive-card" aria-label={attachment.name || "Adaptive Card"}>
      <AdaptiveCardContent elements={attachment.body} onAction={onAction} />
      <div className="message-actions">
        {attachment.actions.flatMap((rawAction, index) => {
          const type = cardText(rawAction.type);
          const title = cardText(rawAction.title);
          if (!title) {
            return [];
          }
          if (type === "Action.OpenUrl") {
            const url = safeHttpUrl(rawAction.url);
            return url ? [
              <a className="action-button" href={url} target="_blank" rel="noreferrer" key={`${title}-${index}`}>
                {title}
              </a>,
            ] : [];
          }
          if (type === "Action.Submit") {
            const data = rawAction.data;
            const value = typeof data === "string"
              ? data
              : data && typeof data === "object"
                ? cardText((data as Record<string, unknown>).text) ||
                  cardText((data as Record<string, unknown>).value)
                : "";
            return value ? [
              <button className="action-button" onClick={() => onAction({ title, value })} key={`${title}-${index}`}>
                {title}
              </button>,
            ] : [];
          }
          return [];
        })}
      </div>
    </section>
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
  const [avatarEnabled, setAvatarEnabled] = React.useState(true);
  const [avatarVisible, setAvatarVisible] = React.useState(false);
  const [voiceConversationActive, setVoiceConversationActive] = React.useState(false);
  const [hostName, setHostName] = React.useState("browser");
  const [agents, setAgents] = React.useState<AgentDefinition[]>([]);
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
  const autoSubmitTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

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
          homeAccountId: context.user?.id,
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
        demoEmailRecipientRef.current = identityContext.loginHint;
        identityRef.current = new MicrosoftIdentity(config, identityContext);
        speechRef.current = new SpeechController(config.defaultAgent, videoRef.current);
        setInputLanguage(config.defaultAgent.locale);
        setAgents(
          config.demoMode && config.demoAgents?.length
            ? config.demoAgents
            : [config.defaultAgent],
        );
        setSelectedAgentId(config.defaultAgent.id);
        setState("ready");
        setStatus("Ready to connect");
      } catch (error) {
        setState("error");
        setStatus(error instanceof Error ? error.message : "Initialization failed.");
      }
    }

    void initialize();
    return () => {
      disposed = true;
      if (autoSubmitTimerRef.current) {
        clearTimeout(autoSubmitTimerRef.current);
      }
      void speechRef.current?.close();
    };
  }, []);

  React.useEffect(() => {
    transcriptRef.current?.scrollTo({
      top: transcriptRef.current.scrollHeight,
      behavior: "smooth",
    });
  }, [messages, interim]);

  async function connect(): Promise<void> {
    const config = configRef.current;
    const identity = identityRef.current;
    if (!config || !identity) {
      return;
    }
    setState("connecting");
    setCanManageCatalog(false);
    setConfigurationUrl("");
    const selectedName =
      agents.find((agent) => agent.id === selectedAgentId)?.displayName ||
      config.defaultAgent.displayName;
    setStatus(`Connecting to ${selectedName}...`);
    try {
      let availableAgents: AgentDefinition[];
      let catalogNotice: string | undefined;
      if (config.demoMode && config.demoAgents?.length) {
        availableAgents = config.demoAgents;
      } else {
        try {
          let userEmail = identity.getBestEffortUserEmail();
          if (config.catalogEnabled && !userEmail) {
            userEmail = await identity.ensureUserIdentity(config.defaultAgent);
          }
          const catalog = await loadAgentCatalog(config, userEmail);
          availableAgents = catalog.agents;
          setCanManageCatalog(catalog.canManageCatalog);
          setConfigurationUrl(catalog.configurationUrl || "");
        } catch (error) {
          if (error instanceof CatalogAccessError) {
            throw error;
          }
          availableAgents = [config.defaultAgent];
          catalogNotice =
            error instanceof Error
              ? `The shared agent catalog is unavailable (${error.message}) Using the configured default agent.`
              : "The shared agent catalog is unavailable. Using the configured default agent.";
        }
      }
      setAgents(availableAgents);
      const selected =
        availableAgents.find((agent) => agent.id === selectedAgentId) ||
        availableAgents.find(
          (agent) =>
            agent.environmentId === config.defaultAgent.environmentId &&
            agent.schemaName === config.defaultAgent.schemaName,
        ) ||
        availableAgents[0];
      setSelectedAgentId(selected.id);
      await speechRef.current?.setAgent(selected);
      setInputLanguage(selected.locale);
      speechRef.current?.setInputLanguage(selected.locale);
      setAvatarVisible(false);
      agentClientRef.current = config.demoMode
        ? new DemoAgentClient(selected, demoEmailRecipientRef.current)
        : new CopilotAgentClient(identity, selected);
      const turn = await agentClientRef.current.connect();
      const welcome =
        turn.messages.length > 0
          ? turn.messages
          : [{ text: selected.welcomeMessage }];
      setMessages([
        ...(catalogNotice ? [createMessage("system", catalogNotice)] : []),
        ...welcome.map((message) => createMessage("agent", message)),
      ]);
      setState("ready");
      setStatus(`${selected.displayName} is ready`);
      await speakMessages(
        welcome
          .map((message) => message.speak || toSpokenText(message.text))
          .filter(Boolean),
      );
    } catch (error) {
      handleError(error, `Could not connect to ${selectedName}.`);
    }
  }

  async function sendMessage(textOverride?: string): Promise<void> {
    const text = (textOverride ?? draft).trim();
    if (!text || !agentClientRef.current || state === "thinking") {
      return;
    }

    await speechRef.current?.stopListening().catch(() => undefined);
    setInterim("");
    setDraft("");
    setMessages((current) => [...current, createMessage("user", text)]);
    setState("thinking");
    setStatus(`${selectedAgent?.displayName || "The agent"} is thinking...`);

    try {
      const turn = await agentClientRef.current.send(text);
      if (turn.messages.length === 0) {
        throw new Error(`${selectedAgent?.displayName || "The agent"} returned no message.`);
      }

      const nextMessages: ChatMessage[] = [];
      const spoken: string[] = [];
      let completed = false;
      const completionPhrase =
        agents.find((agent) => agent.id === selectedAgentId)?.completionPhrase ||
        configRef.current?.defaultAgent.completionPhrase ||
        "";

      for (const reply of turn.messages) {
        completed ||= reply.text
          .toLocaleLowerCase()
          .includes(completionPhrase.toLocaleLowerCase());
        nextMessages.push(createMessage("agent", reply));
        const speech = reply.speak || toSpokenText(reply.text);
        if (speech) {
          spoken.push(speech);
        }
      }

      setMessages((current) => [...current, ...nextMessages]);
      setEmailDraftUrl(
        turn.emailDraft
          ? `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(turn.emailDraft.to)}&subject=${encodeURIComponent(turn.emailDraft.subject)}&body=${encodeURIComponent(turn.emailDraft.body)}`
          : "",
      );
      if (completed) {
        setVoiceConversation(false);
        setState("complete");
        setStatus("Conversation complete");
        await speakMessages(spoken, true);
      } else {
        setState("ready");
        setStatus("Your turn");
        await speakMessages(spoken);
      }
    } catch (error) {
      handleError(error, `${selectedAgent?.displayName || "The agent"} could not complete that turn.`);
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
    setVoiceConversation(true);
    await startListening(true);
  }

  async function stopVoiceConversation(): Promise<void> {
    setVoiceConversation(false);
    await speechRef.current?.stopListening().catch(() => undefined);
    await speechRef.current?.stopSpeaking();
    if (state !== "complete" && state !== "error") {
      setState("ready");
      setStatus("Voice conversation stopped");
    }
  }

  async function speakMessages(texts: string[], preserveCompletion = false): Promise<void> {
    if (!speechRef.current || texts.length === 0) {
      return;
    }
    setState(preserveCompletion ? "complete" : "speaking");
    setStatus(`${selectedAgent?.displayName || "The agent"} is speaking...`);
    try {
      for (const text of texts) {
        await speechRef.current.speak(text, avatarEnabled);
      }
    } catch (error) {
      if (avatarEnabled) {
        setAvatarEnabled(false);
        setMessages((current) => [
          ...current,
          createMessage("system", "Avatar video is unavailable; continuing with audio-only mode."),
        ]);
        for (const text of texts) {
          await speechRef.current.speak(text, false);
        }
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
    setVoiceConversation(false);
    const message = error instanceof Error ? error.message : fallback;
    setState("error");
    setStatus(message || fallback);
    setMessages((current) => [...current, createMessage("system", message || fallback)]);
  }

  async function selectAgent(agentId: string): Promise<void> {
    if (agentId === selectedAgentId) {
      return;
    }
    await speechRef.current?.stopSpeaking();
    await speechRef.current?.stopListening().catch(() => undefined);
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

  const connected = messages.length > 0;
  const busy = state === "connecting" || state === "thinking" || state === "speaking";
  const selectedAgent = agents.find((agent) => agent.id === selectedAgentId) || agents[0];

  function handleAgentAction(action: AgentAction): void {
    if (action.url) {
      window.open(action.url, "_blank", "noopener,noreferrer");
      return;
    }
    if (action.value) {
      void sendMessage(action.value);
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
                disabled={busy || state === "listening"}
              >
                {agents.map((agent) => (
                  <option value={agent.id} key={agent.id}>
                    {agent.displayName}
                  </option>
                ))}
              </select>
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
                <button className="primary-button" onClick={() => void connect()} disabled={busy}>
                  {configRef.current?.demoMode ? "Start demo with " : "Connect to "}
                  {selectedAgent?.displayName || "agent"}
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
                <RichMessageContent message={message} onAction={handleAgentAction} />
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
              onChange={(event) => setDraft(event.target.value)}
              placeholder={
                connected
                  ? "Type a message or use the microphone..."
                  : `Connect to ${selectedAgent?.displayName || "an agent"} first`
              }
              disabled={!connected || busy || state === "complete"}
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
                  disabled={!connected || busy || state === "complete"}
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
                checked={avatarEnabled}
                onChange={(event) => {
                  setAvatarEnabled(event.target.checked);
                  if (!event.target.checked) {
                    setAvatarVisible(false);
                  }
                }}
              />
              <span>Avatar video</span>
            </label>
            <button
              className="text-button"
              onClick={() => void speechRef.current?.stopSpeaking()}
              disabled={state !== "speaking"}
            >
              Stop speaking
            </button>
          </div>

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
