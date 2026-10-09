import * as SpeechSDK from "microsoft-cognitiveservices-speech-sdk";

import type { AgentDefinition, SpeechCredentials } from "./types";

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
  private speaker?: SpeechSDK.SpeakerAudioDestination;
  private finishPlayback?: () => void;
  private avatarSynthesizer?: SpeechSDK.AvatarSynthesizer;
  private peerConnection?: RTCPeerConnection;
  private finalSegments: string[] = [];
  private recognitionLanguage: string;

  constructor(
    private agent: AgentDefinition,
    private readonly video: HTMLVideoElement,
    private readonly getApiAccessToken: () => Promise<string>,
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

  async startListening(
    onInterim: (text: string) => void,
    onFinal: (text: string) => void,
    onError: (message: string) => void,
  ): Promise<void> {
    await this.stopSpeaking();
    await this.stopListening();
    const credentials = await getCredentials(await this.getApiAccessToken());
    const speechConfig = createSpeechConfig(
      credentials,
      this.agent,
      this.recognitionLanguage,
    );
    const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();
    this.finalSegments = [];
    this.recognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);

    this.recognizer.recognizing = (_sender, event) => {
      onInterim([...this.finalSegments, event.result.text].filter(Boolean).join(" "));
    };
    this.recognizer.recognized = (_sender, event) => {
      const text = event.result.text.trim();
      if (text) {
        this.finalSegments.push(text);
        onFinal(this.finalSegments.join(" "));
      }
    };
    this.recognizer.canceled = (_sender, event) => {
      onError(event.errorDetails || "Speech recognition was canceled.");
    };

    await new Promise<void>((resolve, reject) => {
      this.recognizer?.startContinuousRecognitionAsync(resolve, reject);
    });
  }

  async stopListening(): Promise<void> {
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
    await this.stopSpeaking();
    const credentials = await getCredentials(await this.getApiAccessToken());
    const speechConfig = createSpeechConfig(
      credentials,
      this.agent,
      this.recognitionLanguage,
    );

    if (avatarEnabled) {
      await this.speakWithAvatar(text, credentials, speechConfig);
      return;
    }

    speechConfig.speechSynthesisOutputFormat =
      SpeechSDK.SpeechSynthesisOutputFormat.Audio24Khz48KBitRateMonoMp3;
    const speaker = new SpeechSDK.SpeakerAudioDestination();
    this.speaker = speaker;
    const audioConfig = SpeechSDK.AudioConfig.fromSpeakerOutput(speaker);
    this.synthesizer = new SpeechSDK.SpeechSynthesizer(speechConfig, audioConfig);
    await new Promise<void>((resolve, reject) => {
      let synthesized = false;
      let playbackEnded = false;
      const finish = () => {
        if (this.finishPlayback === finish) this.finishPlayback = undefined;
        resolve();
      };
      const fail = (error: Error) => {
        if (this.finishPlayback === finish) this.finishPlayback = undefined;
        reject(error);
      };
      this.finishPlayback = finish;
      speaker.onAudioEnd = () => {
        playbackEnded = true;
        if (synthesized) finish();
      };
      speaker.onAudioStart = () => {
        speaker.internalAudio.addEventListener("error", () => {
          fail(new Error("The browser could not play speech audio."));
        }, { once: true });
      };
      this.synthesizer?.speakTextAsync(
        text,
        (result) => {
          if (result.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
            fail(new Error(result.errorDetails || "Speech synthesis failed."));
            return;
          }
          synthesized = true;
          if (playbackEnded) finish();
        },
        (error) => fail(new Error(error)),
      );
    });
  }

  async stopSpeaking(): Promise<void> {
    if (this.avatarSynthesizer) {
      await this.avatarSynthesizer.stopSpeakingAsync().catch(() => undefined);
    }
    this.speaker?.pause();
    this.speaker = undefined;
    this.finishPlayback?.();
    this.finishPlayback = undefined;
    this.synthesizer?.close();
    this.synthesizer = undefined;
  }

  async close(): Promise<void> {
    await this.stopListening().catch(() => undefined);
    await this.stopSpeaking();
    await this.closeAvatar();
  }

  private async closeAvatar(): Promise<void> {
    if (this.avatarSynthesizer) {
      await this.avatarSynthesizer.close();
      this.avatarSynthesizer = undefined;
    }
    this.peerConnection?.close();
    this.peerConnection = undefined;
    this.video.srcObject = null;
  }

  private async speakWithAvatar(
    text: string,
    credentials: SpeechCredentials,
    speechConfig: SpeechSDK.SpeechConfig,
  ): Promise<void> {
    if (!credentials.relay) {
      throw new Error("Avatar relay credentials are unavailable.");
    }

    if (!this.avatarSynthesizer || !this.peerConnection) {
      this.peerConnection = new RTCPeerConnection({
        iceServers: [
          {
            urls: [credentials.relay.url],
            username: credentials.relay.username,
            credential: credentials.relay.credential,
          },
        ],
      });
      this.peerConnection.addTransceiver("video", { direction: "recvonly" });
      this.peerConnection.addTransceiver("audio", { direction: "recvonly" });
      this.peerConnection.ontrack = (event) => {
        if (event.track.kind === "video" || !this.video.srcObject) {
          this.video.srcObject = event.streams[0];
          void this.video.play();
        }
      };

      const avatarConfig = new SpeechSDK.AvatarConfig(
        this.agent.avatarCharacter,
        this.agent.avatarStyle,
        new SpeechSDK.AvatarVideoFormat(),
      );
      this.avatarSynthesizer = new SpeechSDK.AvatarSynthesizer(speechConfig, avatarConfig);
      const result = await this.avatarSynthesizer.startAvatarAsync(this.peerConnection);
      if (result.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
        throw new Error(result.errorDetails || "The avatar connection failed.");
      }
    }

    const result = await this.avatarSynthesizer.speakTextAsync(text);
    if (result.reason !== SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
      throw new Error(result.errorDetails || "The avatar could not speak.");
    }
  }
}
