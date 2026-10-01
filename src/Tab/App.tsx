import React from "react";
import * as teamsJs from "@microsoft/teams-js";

import { loadAgentCatalog } from "./catalog";
import { CopilotAgentClient, MicrosoftIdentity, toSpokenText } from "./copilot";
import { SpeechController } from "./speech";
import type { AgentDefinition, AppConfig, ChatMessage, TurnState } from "./types";
import "./App.css";

function createMessage(role: ChatMessage["role"], text: string): ChatMessage {
  return {
    id: `${Date.now()}-${crypto.randomUUID()}`,
    role,
    text,
  };
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
  const [hostName, setHostName] = React.useState("browser");
  const [agents, setAgents] = React.useState<AgentDefinition[]>([]);
  const [selectedAgentId, setSelectedAgentId] = React.useState("");
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const configRef = React.useRef<AppConfig>();
  const identityRef = React.useRef<MicrosoftIdentity>();
  const agentClientRef = React.useRef<CopilotAgentClient>();
  const speechRef = React.useRef<SpeechController>();
  const transcriptRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    let disposed = false;

    async function initialize(): Promise<void> {
      let loginHint: string | undefined;
      try {
        await teamsJs.app.initialize();
        const context = await teamsJs.app.getContext();
        loginHint = context.user?.loginHint;
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
        identityRef.current = new MicrosoftIdentity(config, loginHint);
        speechRef.current = new SpeechController(config.defaultAgent, videoRef.current);
        setAgents([config.defaultAgent]);
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
    const selectedName =
      agents.find((agent) => agent.id === selectedAgentId)?.displayName ||
      config.defaultAgent.displayName;
    setStatus(`Signing in to ${selectedName}...`);
    try {
      await identity.signIn();
      const availableAgents = await loadAgentCatalog(config, identity);
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
      agentClientRef.current = new CopilotAgentClient(identity, selected);
      const turn = await agentClientRef.current.connect();
      const welcome =
        turn.messages.length > 0
          ? turn.messages
          : [selected.welcomeMessage];
      setMessages(welcome.map((text) => createMessage("agent", text)));
      setState("ready");
      setStatus(`${selected.displayName} is ready`);
      await speakMessages(welcome);
    } catch (error) {
      handleError(error, `Could not connect to ${selectedName}.`);
    }
  }

  async function sendMessage(): Promise<void> {
    const text = draft.trim();
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
        completed ||= reply
          .toLocaleLowerCase()
          .includes(completionPhrase.toLocaleLowerCase());
        nextMessages.push(createMessage("agent", reply));
        const speech = toSpokenText(reply);
        if (speech) {
          spoken.push(speech);
        }
      }

      setMessages((current) => [...current, ...nextMessages]);
      if (completed) {
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

  async function startListening(): Promise<void> {
    if (!speechRef.current) {
      return;
    }
    try {
      setInterim("");
      setState("listening");
      setStatus("Listening... select I'm done when finished");
      await speechRef.current.startListening(
        (text) => setInterim(text),
        (text) => {
          setInterim(text);
          setDraft(text);
        },
        (message) => handleError(new Error(message), "Speech recognition stopped."),
      );
    } catch (error) {
      handleError(error, "Microphone access failed.");
    }
  }

  async function finishListening(): Promise<void> {
    try {
      await speechRef.current?.stopListening();
      setState("ready");
      setStatus("Review your answer, then send");
      if (interim.trim()) {
        setDraft(interim.trim());
      }
    } catch (error) {
      handleError(error, "Could not stop speech recognition.");
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
  }

  function handleError(error: unknown, fallback: string): void {
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
    agentClientRef.current = undefined;
    setMessages([]);
    setDraft("");
    setInterim("");
    setSelectedAgentId(agentId);
    const selected = agents.find((agent) => agent.id === agentId);
    if (selected) {
      await speechRef.current?.setAgent(selected);
    }
    setState("ready");
    setStatus(selected ? `Ready to connect to ${selected.displayName}` : "Ready to connect");
  }

  const connected = messages.length > 0;
  const busy = state === "connecting" || state === "thinking" || state === "speaking";
  const selectedAgent = agents.find((agent) => agent.id === selectedAgentId) || agents[0];

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Voice-enabled Copilot Studio</p>
          <h1>{selectedAgent ? `Meet ${selectedAgent.displayName}` : "Voice interviews"}</h1>
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
          </div>

          <div className="transcript" ref={transcriptRef} aria-live="polite">
            {!connected && (
              <div className="welcome-card">
                <div className="pat-mark">{selectedAgent?.displayName.charAt(0) || "A"}</div>
                <h3>Ready when you are</h3>
                <p>
                  Sign in with your work account. You can type or use your microphone, and
                  you can turn avatar video off at any time.
                </p>
                <button className="primary-button" onClick={() => void connect()} disabled={busy}>
                  Connect to {selectedAgent?.displayName || "agent"}
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
                <p>{message.text}</p>
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
              {state === "listening" ? (
                <button className="secondary-button listening" onClick={() => void finishListening()}>
                  I'm done
                </button>
              ) : (
                <button
                  className="secondary-button"
                  onClick={() => void startListening()}
                  disabled={!connected || busy || state === "complete"}
                >
                  Use microphone
                </button>
              )}
              <button
                className="primary-button"
                onClick={() => void sendMessage()}
                disabled={!draft.trim() || !connected || busy || state === "complete"}
              >
                Send
              </button>
            </div>
          </div>
        </div>

        <aside className="avatar-panel">
          <div className="avatar-stage">
            <video
              ref={videoRef}
              autoPlay
              playsInline
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
                onChange={(event) => setAvatarEnabled(event.target.checked)}
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
                {selectedAgent.avatarStyle})
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
