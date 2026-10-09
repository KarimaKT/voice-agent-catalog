import * as SpeechSDK from "microsoft-cognitiveservices-speech-sdk";

import type { AgentDefinition, AvatarIdleTimeoutSeconds, SpeechCredentials } from "./types";

async function getCredentials(accessToken: string): Promise<SpeechCredentials> {
  const response = await fetch("/api/speech/token", {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    const body = (await response.json().catch(() => undefined)) as { error?: string } | undefined;
    throw new Error(body?.error || "Unable to obtain Speech credentials.");
  }
  return (await response.json()) as SpeechCredentials;
}

function createSpeechConfig(
  credentials: SpeechCredentials,
  agent: AgentDefinition,
  recognitionLanguage: string,
): SpeechSDK.SpeechConfig {
  const speechConfig = SpeechSDK.SpeechConfig.fromAuthorizationToken(
    credentials.token,
    credentials.region,
  );
  speechConfig.speechRecognitionLanguage = recognitionLanguage;
  speechConfig.speechSynthesisLanguage = agent.locale;
  speechConfig.speechSynthesisVoiceName = agent.voiceName;
  return speechConfig;
}

export class SpeechController {
  private recognizer?: SpeechSDK.SpeechRecognizer;
  private synthesizer?: SpeechSDK.SpeechSynthesizer;
  private finishPlayback?: () => void;
  private avatarSynthesizer?: SpeechSDK.AvatarSynthesizer;
  private peerConnection?: RTCPeerConnection;
  private finalSegments: string[] = [];
  private recognitionLanguage: string;
  private listeningGeneration = 0;
  private speakingGeneration = 0;
  private conversationActive = false;
  private avatarGeneration = 0;
  private cancelAvatar?: () => void;
  private avatarIdleSeconds: AvatarIdleTimeoutSeconds = 34;
  private avatarIdleTimer?: ReturnType<typeof setTimeout>;
  private avatarBusy = false;
  private agentWorking = false;

  constructor(
    private agent: AgentDefinition,
    private readonly video: HTMLVideoElement,
    private readonly getApiAccessToken: () => Promise<string>,
    private readonly onAvatarNotice: (message: string) => void = () => {},
  ) {
    this.recognitionLanguage = agent.locale;
  }

  async setAgent(agent: AgentDefinition): Promise<void> {
    const profileChanged =
      this.agent.locale !== agent.locale ||
      this.agent.voiceName !== agent.voiceName ||
      this.agent.avatarCharacter !== agent.avatarCharacter ||
      this.agent.avatarStyle !== agent.avatarStyle;
    if (profileChanged) {
      await this.closeAvatar();
    }
    this.agent = agent;
  }

  setInputLanguage(locale: string): void {
    this.recognitionLanguage = locale;
  }

  async setConversationActive(active: boolean): Promise<void> {
    this.conversationActive = active;
    if (!active) await this.close();
  }

  setAvatarIdleTimeout(seconds: AvatarIdleTimeoutSeconds): void {
    if (seconds !== 15 && seconds !== 34 && seconds !== 45) {
      throw new Error("Avatar idle timeout must be 15, 34 or 45 seconds.");
    }
    this.avatarIdleSeconds = seconds;
    this.resetAvatarIdleTimer();
  }

  setAgentWorking(working: boolean): void {
    this.agentWorking = working;
    this.resetAvatarIdleTimer();
  }

  noteUserActivity(): void {
    this.resetAvatarIdleTimer();
  }

  private resetAvatarIdleTimer(): void {
    if (this.avatarIdleTimer !== undefined) clearTimeout(this.avatarIdleTimer);
    this.avatarIdleTimer = undefined;
    if (!this.conversationActive || !this.avatarSynthesizer ||
        this.avatarBusy || this.agentWorking) return;
    const seconds = this.avatarIdleSeconds;
    const timer = setTimeout(() => {
      if (this.avatarIdleTimer !== timer) return;
      this.avatarIdleTimer = undefined;
      if (this.agentWorking || this.avatarBusy) {
        this.resetAvatarIdleTimer();
        return;
      }
      void this.closeAvatar().then(() => {
        if (!this.avatarSynthesizer) {
          this.onAvatarNotice(`Avatar disconnected after ${seconds} seconds idle to save cost. It will reopen for the next reply.`);
        }
      }).catch((error: unknown) => {
        this.onAvatarNotice(error instanceof Error ? `Avatar cleanup: ${error.message}` : "Avatar cleanup failed.");
      });
    }, seconds * 1000);
    this.avatarIdleTimer = timer;
  }

  async stopAvatar(): Promise<void> {
    this.speakingGeneration++;
    await this.closeAvatar();
  }

  async startListening(
    onInterim: (text: string) => void,
    onFinal: (text: string) => void,
    onError: (message: string) => void,
  ): Promise<void> {
    this.noteUserActivity();
    const generation = ++this.listeningGeneration;
    await this.stopSpeaking();
    if (generation !== this.listeningGeneration) return;
    await this.closeRecognizer();
    const credentials = await getCredentials(await this.getApiAccessToken());
    if (generation !== this.listeningGeneration) return;
    const speechConfig = createSpeechConfig(
      credentials,
      this.agent,
      this.recognitionLanguage,
    );
    const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();
    this.finalSegments = [];
    this.recognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);

    this.recognizer.recognizing = (_sender, event) => {
      if (generation !== this.listeningGeneration) return;
      this.noteUserActivity();
      onInterim([...this.finalSegments, event.result.text].filter(Boolean).join(" "));
    };
    this.recognizer.recognized = (_sender, event) => {
      if (generation !== this.listeningGeneration) return;
      const text = event.result.text.trim();
      if (text) {
        this.noteUserActivity();
        this.finalSegments.push(text);
        onFinal(this.finalSegments.join(" "));
      }
    };
    this.recognizer.canceled = (_sender, event) => {
      if (generation !== this.listeningGeneration) return;
      onError(event.errorDetails || "Speech recognition was canceled.");
    };

    await new Promise<void>((resolve, reject) => {
      this.recognizer?.startContinuousRecognitionAsync(resolve, reject);
    });
  }

  async stopListening(): Promise<void> {
    this.listeningGeneration++;
    await this.closeRecognizer();
  }

  private async closeRecognizer(): Promise<void> {
    if (!this.recognizer) {
      return;
    }
    const recognizer = this.recognizer;
    this.recognizer = undefined;
    await new Promise<void>((resolve, reject) => {
      recognizer.stopContinuousRecognitionAsync(
        () => {
          recognizer.close();
          resolve();
        },
        reject,
      );
    });
  }

  async speak(text: string, avatarEnabled: boolean): Promise<void> {
    if (!text) {
      return;
    }
    if (avatarEnabled && !this.conversationActive) {
      throw new Error("Avatar requires an active Copilot Studio conversation.");
    }
    const stopping = this.stopSpeaking();
    const generation = this.speakingGeneration;
    await stopping;
    const credentials = await getCredentials(await this.getApiAccessToken());
    if (generation !== this.speakingGeneration) return;
    const speechConfig = createSpeechConfig(
      credentials,
      this.agent,
      this.recognitionLanguage,
    );

    if (avatarEnabled) {
      if (!this.conversationActive) {
        throw new Error("Avatar requires an active Copilot Studio conversation.");
      }
      await this.speakWithAvatar(text, credentials, speechConfig, generation);
      return;
    }

    speechConfig.speechSynthesisOutputFormat =
      SpeechSDK.SpeechSynthesisOutputFormat.Riff16Khz16BitMonoPcm;
    this.synthesizer = new SpeechSDK.SpeechSynthesizer(speechConfig, null);
    await new Promise<void>((resolve, reject) => {
      let audio: HTMLAudioElement | undefined;
      let objectUrl: string | undefined;
      let settled = false;
      const settle = (error?: Error) => {
        if (settled) return;
        settled = true;
        audio?.pause();
        if (objectUrl) URL.revokeObjectURL(objectUrl);
        if (this.finishPlayback === finish) this.finishPlayback = undefined;
        if (error) reject(error);
        else resolve();
      };
      const finish = () => settle();
      this.finishPlayback = finish;
      this.synthesizer?.speakTextAsync(
        text,
        (result) => {
          if (settled) return;
          if (result.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
            settle(new Error(result.errorDetails || "Speech synthesis failed."));
            return;
          }
          if (!result.audioData.byteLength) {
            settle(new Error("Speech synthesis returned no audio."));
            return;
          }
          objectUrl = URL.createObjectURL(new Blob([result.audioData], { type: "audio/wav" }));
          audio = new Audio(objectUrl);
          audio.onended = finish;
          audio.onerror = () => settle(new Error("The browser could not play speech audio."));
          void audio.play().catch((error: unknown) => {
            settle(new Error(error instanceof Error ? error.message : "Speech playback could not start."));
          });
        },
        (error) => settle(new Error(error)),
      );
    });
  }

  async stopSpeaking(): Promise<void> {
    this.speakingGeneration++;
    this.cancelAvatar?.();
    if (this.avatarBusy && this.avatarSynthesizer) {
      await this.avatarSynthesizer.stopSpeakingAsync();
    }
    this.finishPlayback?.();
    this.finishPlayback = undefined;
    this.synthesizer?.close();
    this.synthesizer = undefined;
  }

  async close(): Promise<void> {
    this.conversationActive = false;
    this.speakingGeneration++;
    this.agentWorking = false;
    const results = await Promise.allSettled([
      this.closeAvatar(), this.stopSpeaking(), this.stopListening(),
    ]);
    for (const result of results) {
      if (result.status === "rejected") throw result.reason;
    }
  }

  private async closeAvatar(): Promise<void> {
    this.avatarGeneration++;
    if (this.avatarIdleTimer !== undefined) clearTimeout(this.avatarIdleTimer);
    this.avatarIdleTimer = undefined;
    this.avatarBusy = false;
    this.cancelAvatar?.();
    this.cancelAvatar = undefined;
    const avatar = this.avatarSynthesizer;
    const peer = this.peerConnection;
    this.avatarSynthesizer = undefined;
    this.peerConnection = undefined;
    peer?.close();
    const stream = this.video.srcObject;
    if (stream instanceof MediaStream) {
      for (const track of stream.getTracks()) track.stop();
    }
    this.video.srcObject = null;
    if (avatar) await avatar.close();
  }

  private async speakWithAvatar(
    text: string,
    credentials: SpeechCredentials,
    speechConfig: SpeechSDK.SpeechConfig,
    generation: number,
  ): Promise<void> {
    const avatarGeneration = this.avatarGeneration;
    this.avatarBusy = true;
    this.resetAvatarIdleTimer();
    let succeeded = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try {
      let peer = this.peerConnection;
      let avatar = this.avatarSynthesizer;
      const needsStartup = !peer || !avatar;
      if (!peer || !avatar) {
        if (!credentials.relay) throw new Error("Avatar relay credentials are unavailable.");
        peer = new RTCPeerConnection({
          iceServers: [{
            urls: [credentials.relay.url],
            username: credentials.relay.username,
            credential: credentials.relay.credential,
          }],
        });
        this.peerConnection = peer;
        peer.addTransceiver("video", { direction: "recvonly" });
        peer.addTransceiver("audio", { direction: "recvonly" });
        const avatarConfig = new SpeechSDK.AvatarConfig(
          this.agent.avatarCharacter,
          this.agent.avatarStyle,
          new SpeechSDK.AvatarVideoFormat(),
        );
        avatar = new SpeechSDK.AvatarSynthesizer(speechConfig, avatarConfig);
        this.avatarSynthesizer = avatar;
      }
      let failMedia!: (error: Error) => void;
      peer.ontrack = (event) => {
        if (avatarGeneration !== this.avatarGeneration || !this.conversationActive) {
          event.track.stop();
          return;
        }
        if (event.track.kind === "video" || !this.video.srcObject) {
          this.video.srcObject = event.streams[0];
          void this.video.play().catch((error: unknown) => {
            failMedia(new Error(error instanceof Error ? error.message : "Avatar playback failed."));
          });
        }
      };

      const canceled = {
        reason: SpeechSDK.ResultReason.SynthesizingAudioCompleted,
        errorDetails: "",
      };
      const cancellation = new Promise<typeof canceled>((resolve) => {
        this.cancelAvatar = () => resolve(canceled);
      });
      const connectionFailure = new Promise<never>((_resolve, reject) => {
        failMedia = reject;
        peer.onconnectionstatechange = () => {
          if (this.peerConnection?.connectionState === "failed" ||
              this.peerConnection?.connectionState === "disconnected") {
            const error = new Error("Avatar media connection was lost.");
            reject(error);
            void this.closeAvatar().then(() => {
              this.onAvatarNotice(error.message);
            }).catch((cleanupError: unknown) => {
              this.onAvatarNotice(cleanupError instanceof Error ? cleanupError.message : "Avatar cleanup failed.");
            });
          }
        };
      });
      if (needsStartup) {
        const result = await Promise.race([
          avatar.startAvatarAsync(peer),
          cancellation,
          connectionFailure,
          new Promise<never>((_resolve, reject) => {
            timeout = setTimeout(() => reject(new Error("Avatar startup timed out.")), 30_000);
          }),
        ]);
        if (timeout !== undefined) clearTimeout(timeout);
        if (result.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
          throw new Error(result.errorDetails || "The avatar connection failed.");
        }
      }
      if (generation !== this.speakingGeneration ||
          avatarGeneration !== this.avatarGeneration || !this.conversationActive) return;
      const spoken = await Promise.race([
        avatar.speakTextAsync(text),
        cancellation,
        connectionFailure,
        new Promise<never>((_resolve, reject) => {
          timeout = setTimeout(() => reject(new Error("Avatar speech timed out.")), 120_000);
        }),
      ]);
      if (spoken.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
        throw new Error(spoken.errorDetails || "The avatar could not speak.");
      }
      succeeded = generation === this.speakingGeneration &&
        avatarGeneration === this.avatarGeneration && this.conversationActive;
    } finally {
      if (timeout) clearTimeout(timeout);
      if (avatarGeneration === this.avatarGeneration) {
        this.avatarBusy = false;
        this.cancelAvatar = undefined;
        if (succeeded) this.resetAvatarIdleTimer();
        else await this.closeAvatar();
      }
    }
  }
}
