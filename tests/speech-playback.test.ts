import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import { runInNewContext } from "node:vm";

function speechHarness() {
  type Result = { reason: number; audioData: ArrayBuffer; errorDetails?: string };
  const players: Player[] = [];
  const synthesizers: Synthesizer[] = [];
  const config: Record<string, unknown> = {};
  const revoked: string[] = [];
  const avatars: Avatar[] = [];
  const peers: Peer[] = [];
  const timers = new Map<number, { callback: () => void; delay: number }>();
  let nextTimer = 0;
  let delayAvatarStart = false;
  let avatarFailure = false;
  let closeFailure = false;
  class Stream {
    track = { stopped: false, stop() { this.stopped = true; } };
    getTracks() { return [this.track]; }
  }
  const video = { srcObject: null as Stream | null, play: async () => {} };
  class Peer {
    closed = false;
    connectionState = "connected";
    ontrack?: (event: { track: { kind: string; stop(): void }; streams: Stream[] }) => void;
    onconnectionstatechange?: () => void;
    constructor() { peers.push(this); }
    addTransceiver() {}
    close() { this.closed = true; }
  }
  class Avatar {
    closed = false;
    spoken = 0;
    finish?: () => void;
    start?: () => void;
    constructor() { avatars.push(this); }
    startAvatarAsync() {
      if (avatarFailure) return Promise.reject(new Error("Avatar startup failed"));
      if (delayAvatarStart) return new Promise<{ reason: number }>((resolve) => {
        this.start = () => resolve({ reason: 1 });
      });
      return Promise.resolve({ reason: 1 });
    }
    speakTextAsync() {
      this.spoken++;
      return new Promise<{ reason: number }>((resolve) => {
        this.finish = () => resolve({ reason: 1 });
      });
    }
    async close() {
      this.closed = true;
      if (closeFailure) throw new Error("Avatar close failed");
    }
  }
  let playError: Error | undefined;
  let credentialsGate: Promise<void> | undefined;
  let credentialRequests = 0;
  class Player {
    paused = false;
    onended?: () => void;
    onerror?: () => void;
    constructor() { players.push(this); }
    play() { return playError ? Promise.reject(playError) : Promise.resolve(); }
    pause() { this.paused = true; }
  }
  class Synthesizer {
    closed = false;
    complete?: (result: Result) => void;
    constructor(_config: unknown, output: unknown) {
      assert.equal(output, null, "Native playback owns output instead of the SDK's streaming player");
      synthesizers.push(this);
    }
    speakTextAsync(_text: string, complete: (result: Result) => void) { this.complete = complete; }
    close() { this.closed = true; }
  }
  const source = readFileSync(new URL("../src/Tab/speech.ts", import.meta.url), "utf8");
  const script = stripTypeScriptTypes(source, { mode: "transform" })
    .replace(/^import[\s\S]*?;\r?$/gm, "")
    .replace(/^export /gm, "");
  const Controller = runInNewContext(`${script}\nSpeechController;`, {
    Error, Blob, Audio: Player, MediaStream: Stream, RTCPeerConnection: Peer,
    setTimeout: (callback: () => void, delay: number) => {
      const id = ++nextTimer;
      timers.set(id, { callback, delay });
      return id;
    },
    clearTimeout: (id: number) => timers.delete(id),
    URL: { createObjectURL: () => "blob:test-audio", revokeObjectURL: (url: string) => revoked.push(url) },
    fetch: async () => {
      credentialRequests++;
      if (credentialsGate) await credentialsGate;
      return { ok: true, json: async () => ({
        token: "test-token", region: "test-region",
        relay: { url: "turn:test", username: "test", credential: "test" },
      }) };
    },
    SpeechSDK: {
      SpeechConfig: { fromAuthorizationToken: () => config },
      SpeechSynthesisOutputFormat: { Riff16Khz16BitMonoPcm: 1 },
      SpeechSynthesizer: Synthesizer,
      AvatarSynthesizer: Avatar,
      AvatarConfig: class {},
      AvatarVideoFormat: class {},
      ResultReason: { SynthesizingAudioCompleted: 1 },
    },
  }) as new (...args: unknown[]) => {
    speak(text: string, avatar: boolean): Promise<void>;
    stopSpeaking(): Promise<void>;
    setInputLanguage(locale: string): void;
    startListening(...callbacks: unknown[]): Promise<void>;
    stopListening(): Promise<void>;
    setConversationActive(active: boolean): Promise<void>;
    stopAvatar(): Promise<void>;
    close(): Promise<void>;
  };
  return {
    controller: new Controller({ locale: "en-US", voiceName: "test-voice" }, video, async () => "api-token"),
    players, synthesizers, config, revoked,
    avatars, peers, timers, video, Stream,
    delayAvatar: () => { delayAvatarStart = true; },
    failAvatar: () => { avatarFailure = true; },
    failClose: () => { closeFailure = true; },
    avatarReady: async () => {
      for (let index = 0; index < 30 && !avatars[0]?.finish; index++) {
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      assert.ok(avatars[0]?.finish, "Avatar speech started");
    },
    failPlayback: () => { playError = new Error("Autoplay denied"); },
    result: { reason: 1, audioData: new ArrayBuffer(16) },
    delayCredentials: () => {
      let release!: () => void;
      credentialsGate = new Promise<void>((resolve) => { release = resolve; });
      return release;
    },
    credentialsRequested: async () => {
      for (let index = 0; index < 20 && !credentialRequests; index++) {
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      assert.equal(credentialRequests, 1);
    },
    ready: async () => {
      for (let index = 0; index < 20 && !synthesizers[0]?.complete; index++) {
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      assert.ok(synthesizers[0]?.complete, "Synthesis started");
    },
  };
}

test("audio-only speech waits for native playback end rather than synthesis completion", async () => {
  const harness = speechHarness();
  let finished = false;
  const speech = harness.controller.speak("A reply", false).then(() => { finished = true; });
  await harness.ready();
  assert.equal(harness.config.speechSynthesisOutputFormat, 1);
  harness.synthesizers[0].complete!(harness.result);
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.equal(finished, false);
  harness.players[0].onended!();
  await speech;
  assert.equal(finished, true);
  assert.deepEqual(harness.revoked, ["blob:test-audio"]);
});

test("avatar cannot start without an active Copilot Studio session", async () => {
  const harness = speechHarness();
  await assert.rejects(harness.controller.speak("A reply", true), /active Copilot Studio conversation/);
  assert.equal(harness.avatars.length, 0);
  assert.equal(harness.peers.length, 0);
});

test("avatar closes and stops its media after every reply, leaving no paid idle session", async () => {
  const harness = speechHarness();
  await harness.controller.setConversationActive(true);
  const speech = harness.controller.speak("A reply", true);
  await harness.avatarReady();
  const stream = new harness.Stream();
  harness.peers[0].ontrack!({ track: { kind: "video", stop() {} }, streams: [stream] });
  harness.avatars[0].finish!();
  await speech;
  assert.equal(harness.avatars[0].closed, true);
  assert.equal(harness.peers[0].closed, true);
  assert.equal(stream.track.stopped, true);
  assert.equal(harness.video.srcObject, null);
  assert.equal(harness.timers.size, 0);
});

test("ending a session or turning avatar off immediately cancels a running avatar", async () => {
  for (const sessionEnded of [true, false]) {
    const harness = speechHarness();
    await harness.controller.setConversationActive(true);
    const speech = harness.controller.speak("A reply", true);
    await harness.avatarReady();
    if (sessionEnded) await harness.controller.setConversationActive(false);
    else await harness.controller.stopAvatar();
    await speech;
    assert.equal(harness.avatars[0].closed, true);
    assert.equal(harness.peers[0].closed, true);
    assert.equal(harness.timers.size, 0);
    if (sessionEnded) {
      await assert.rejects(harness.controller.speak("Late reply", true), /active Copilot Studio conversation/);
    }
  }
});

test("late avatar startup and tracks cannot resurrect an ended session", async () => {
  const harness = speechHarness();
  harness.delayAvatar();
  await harness.controller.setConversationActive(true);
  const speech = harness.controller.speak("A reply", true);
  await harness.credentialsRequested();
  for (let index = 0; index < 20 && !harness.avatars[0]?.start; index++) {
    await new Promise<void>((resolve) => setImmediate(resolve));
  }
  await harness.controller.setConversationActive(false);
  harness.avatars[0].start!();
  let stopped = false;
  harness.peers[0].ontrack!({
    track: { kind: "video", stop() { stopped = true; } },
    streams: [new harness.Stream()],
  });
  await speech;
  assert.equal(harness.avatars[0].spoken, 0);
  assert.equal(stopped, true);
  assert.equal(harness.video.srcObject, null);
  assert.equal(harness.timers.size, 0);
});

test("startup failure and media disconnect release avatar resources", async () => {
  for (const startupFails of [true, false]) {
    const harness = speechHarness();
    await harness.controller.setConversationActive(true);
    if (startupFails) harness.failAvatar();
    const rejected = assert.rejects(harness.controller.speak("A reply", true), /startup failed|connection was lost/);
    if (!startupFails) {
      await harness.avatarReady();
      harness.peers[0].connectionState = "disconnected";
      harness.peers[0].onconnectionstatechange!();
    }
    await rejected;
    assert.equal(harness.avatars[0].closed, true);
    assert.equal(harness.peers[0].closed, true);
    assert.equal(harness.video.srcObject, null);
    assert.equal(harness.timers.size, 0);
  }
});

test("avatar startup and speech watchdogs enforce exact deadlines and release resources", async () => {
  for (const startup of [true, false]) {
    const harness = speechHarness();
    if (startup) harness.delayAvatar();
    await harness.controller.setConversationActive(true);
    const rejected = assert.rejects(harness.controller.speak("A reply", true), /timed out/);
    if (!startup) await harness.avatarReady();
    else await harness.credentialsRequested();
    for (let index = 0; index < 20 && !harness.timers.size; index++) {
      await new Promise<void>((resolve) => setImmediate(resolve));
    }
    const timer = [...harness.timers.values()][0];
    assert.equal(timer.delay, startup ? 30_000 : 120_000);
    timer.callback();
    await rejected;
    assert.equal(harness.avatars[0].closed, true);
    assert.equal(harness.peers[0].closed, true);
    assert.equal(harness.timers.size, 0);
  }
});

test("SDK close errors surface after local peer and track cleanup", async () => {
  const harness = speechHarness();
  await harness.controller.setConversationActive(true);
  const speech = harness.controller.speak("A reply", true);
  const rejected = assert.rejects(speech, /Avatar close failed/);
  await harness.avatarReady();
  harness.failClose();
  harness.avatars[0].finish!();
  await rejected;
  assert.equal(harness.peers[0].closed, true);
  assert.equal(harness.video.srcObject, null);
});

test("canceled and empty synthesis results surface explicit errors", async () => {
  for (const result of [
    { reason: 0, audioData: new ArrayBuffer(0), errorDetails: "Speech service unavailable" },
    { reason: 1, audioData: new ArrayBuffer(0) },
  ]) {
    const harness = speechHarness();
    const rejected = assert.rejects(harness.controller.speak("A reply", false), /Speech service unavailable|returned no audio/);
    await harness.ready();
    harness.synthesizers[0].complete!(result);
    await rejected;
    assert.equal(harness.players.length, 0);
  }
});

test("Stop pauses output, revokes its URL and releases the pending playback wait", async () => {
  const harness = speechHarness();
  const speech = harness.controller.speak("A reply", false);
  await harness.ready();
  harness.synthesizers[0].complete!(harness.result);
  await harness.controller.stopSpeaking();
  await speech;
  assert.equal(harness.players[0].paused, true);
  assert.equal(harness.synthesizers[0].closed, true);
  assert.deepEqual(harness.revoked, ["blob:test-audio"]);
});

test("Stop during synthesis prevents late audio from starting", async () => {
  const harness = speechHarness();
  const speech = harness.controller.speak("A reply", false);
  await harness.ready();
  await harness.controller.stopSpeaking();
  harness.synthesizers[0].complete!(harness.result);
  await speech;
  assert.equal(harness.players.length, 0);
});

test("native audio errors and autoplay rejection are visible and clean up output", async () => {
  for (const autoplay of [false, true]) {
    const harness = speechHarness();
    if (autoplay) harness.failPlayback();
    const rejected = assert.rejects(harness.controller.speak("A reply", false), /could not play speech audio|Autoplay denied/);
    await harness.ready();
    harness.synthesizers[0].complete!(harness.result);
    if (!autoplay) harness.players[0].onerror!();
    await rejected;
    assert.deepEqual(harness.revoked, ["blob:test-audio"]);
  }
});

test("recognition language stays independent of catalog synthesis language and voice", async () => {
  const harness = speechHarness();
  harness.controller.setInputLanguage("es-ES");
  const speech = harness.controller.speak("A reply", false);
  await harness.ready();
  assert.equal(harness.config.speechRecognitionLanguage, "es-ES");
  assert.equal(harness.config.speechSynthesisLanguage, "en-US");
  assert.equal(harness.config.speechSynthesisVoiceName, "test-voice");
  harness.synthesizers[0].complete!(harness.result);
  harness.players[0].onended!();
  await speech;
});

test("Stop during credential acquisition prevents late synthesis from starting", async () => {
  const harness = speechHarness();
  const release = harness.delayCredentials();
  const speech = harness.controller.speak("A reply", false);
  await harness.credentialsRequested();
  await harness.controller.stopSpeaking();
  release();
  await speech;
  assert.equal(harness.synthesizers.length, 0);
  assert.equal(harness.players.length, 0);
});

test("Stop during credential acquisition prevents late microphone acquisition", async () => {
  const harness = speechHarness();
  const release = harness.delayCredentials();
  const listening = harness.controller.startListening(() => {}, () => {}, () => {});
  await harness.credentialsRequested();
  await harness.controller.stopListening();
  release();
  await listening;
});
