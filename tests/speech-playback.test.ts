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
  let playError: Error | undefined;
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
    Error, Blob, Audio: Player,
    URL: { createObjectURL: () => "blob:test-audio", revokeObjectURL: (url: string) => revoked.push(url) },
    fetch: async () => ({ ok: true, json: async () => ({ token: "test-token", region: "test-region" }) }),
    SpeechSDK: {
      SpeechConfig: { fromAuthorizationToken: () => config },
      SpeechSynthesisOutputFormat: { Riff16Khz16BitMonoPcm: 1 },
      SpeechSynthesizer: Synthesizer,
      ResultReason: { SynthesizingAudioCompleted: 1 },
    },
  }) as new (...args: unknown[]) => {
    speak(text: string, avatar: boolean): Promise<void>;
    stopSpeaking(): Promise<void>;
    setInputLanguage(locale: string): void;
  };
  return {
    controller: new Controller({ locale: "en-US", voiceName: "test-voice" }, {}, async () => "api-token"),
    players, synthesizers, config, revoked,
    failPlayback: () => { playError = new Error("Autoplay denied"); },
    result: { reason: 1, audioData: new ArrayBuffer(16) },
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
