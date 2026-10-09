import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import { runInNewContext } from "node:vm";

function speechHarness() {
  type Result = { reason: number; errorDetails?: string };
  const players: Player[] = [];
  const synthesizers: Synthesizer[] = [];
  class Player {
    paused = false;
    onAudioEnd?: () => void;
    onAudioStart?: () => void;
    internalAudio = new EventTarget();
    constructor() { players.push(this); }
    pause() { this.paused = true; }
  }
  class Synthesizer {
    closed = false;
    complete?: (result: Result) => void;
    constructor() { synthesizers.push(this); }
    speakTextAsync(_text: string, complete: (result: Result) => void) {
      this.complete = complete;
    }
    close() { this.closed = true; }
  }
  const source = readFileSync(new URL("../src/Tab/speech.ts", import.meta.url), "utf8");
  const script = stripTypeScriptTypes(source, { mode: "transform" })
    .replace(/^import[\s\S]*?;\r?$/gm, "")
    .replace(/^export /gm, "");
  const Controller = runInNewContext(`${script}\nSpeechController;`, {
    Error,
    fetch: async () => ({ ok: true, json: async () => ({ token: "test-token", region: "test-region" }) }),
    SpeechSDK: {
      SpeechConfig: { fromAuthorizationToken: () => ({}) },
      SpeakerAudioDestination: Player,
      AudioConfig: { fromSpeakerOutput: (player: Player) => player },
      SpeechSynthesizer: Synthesizer,
      ResultReason: { SynthesizingAudioCompleted: 1 },
    },
  }) as new (...args: unknown[]) => {
    speak(text: string, avatar: boolean): Promise<void>;
    stopSpeaking(): Promise<void>;
  };
  return {
    controller: new Controller({ locale: "en-US", voiceName: "test-voice" }, {}, async () => "api-token"),
    players,
    synthesizers,
    ready: async () => {
      for (let index = 0; index < 20 && !synthesizers[0]?.complete; index++) {
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      assert.ok(synthesizers[0]?.complete, "Synthesis started");
    },
  };
}

test("audio-only speech waits for playback end, not just service synthesis completion", async () => {
  const harness = speechHarness();
  let finished = false;
  const speech = harness.controller.speak("A reply", false).then(() => { finished = true; });
  await harness.ready();
  harness.synthesizers[0].complete!({ reason: 1 });
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.equal(finished, false, "Listening must remain paused while the speaker plays");
  harness.players[0].onAudioEnd!();
  await speech;
  assert.equal(finished, true);
});

test("playback ending before the synthesis callback still waits for successful synthesis", async () => {
  const harness = speechHarness();
  let finished = false;
  const speech = harness.controller.speak("A reply", false).then(() => { finished = true; });
  await harness.ready();
  harness.players[0].onAudioEnd!();
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.equal(finished, false);
  harness.synthesizers[0].complete!({ reason: 1 });
  await speech;
  assert.equal(finished, true);
});

test("a canceled synthesis result surfaces its error instead of appearing successful", async () => {
  const harness = speechHarness();
  const speech = harness.controller.speak("A reply", false);
  const rejected = assert.rejects(speech, /Speech service unavailable/);
  await harness.ready();
  harness.synthesizers[0].complete!({ reason: 0, errorDetails: "Speech service unavailable" });
  await rejected;
});

test("Stop speaking pauses the output and releases the pending playback wait", async () => {
  const harness = speechHarness();
  const speech = harness.controller.speak("A reply", false);
  await harness.ready();
  harness.synthesizers[0].complete!({ reason: 1 });
  await harness.controller.stopSpeaking();
  await speech;
  assert.equal(harness.players[0].paused, true);
  assert.equal(harness.synthesizers[0].closed, true);
});

test("browser audio errors are visible to the conversation controller", async () => {
  const harness = speechHarness();
  const speech = harness.controller.speak("A reply", false);
  const rejected = assert.rejects(speech, /browser could not play speech audio/);
  await harness.ready();
  harness.players[0].onAudioStart!();
  harness.players[0].internalAudio.dispatchEvent(new Event("error"));
  await rejected;
});
