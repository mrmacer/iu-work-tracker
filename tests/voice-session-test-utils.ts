// Shared test helpers for Patch 7.2 (Voice Intelligence working session). This environment's
// global jsdom Storage is unreliable under Node 25 (Storage.prototype spies don't intercept it,
// and `spyOn(window.sessionStorage, "setItem")` fails), so tests use a small in-memory fake —
// passed straight to the helper/component via the `storage` seam, or installed on `window` with
// installWindowStorage() when the DEFAULT browser wiring itself is under test.
import type { VoiceSessionStorage } from "../lib/voice-intelligence-session";

export type FakeStorage = VoiceSessionStorage & {
  data: Map<string, string>;
  setItemCalls: [string, string][];
  removeItemCalls: string[];
};

export function createFakeStorage(): FakeStorage {
  const data = new Map<string, string>();
  const setItemCalls: [string, string][] = [];
  const removeItemCalls: string[] = [];
  return {
    data,
    setItemCalls,
    removeItemCalls,
    getItem: (key) => data.get(key) ?? null,
    setItem: (key, value) => {
      setItemCalls.push([key, value]);
      data.set(key, value);
    },
    removeItem: (key) => {
      removeItemCalls.push(key);
      data.delete(key);
    },
  };
}

/** Replaces window.localStorage / window.sessionStorage with the fake; returns a restore function. */
export function installWindowStorage(name: "localStorage" | "sessionStorage", fake: VoiceSessionStorage): () => void {
  const original = Object.getOwnPropertyDescriptor(window, name);
  Object.defineProperty(window, name, { value: fake, configurable: true });
  return () => {
    if (original) Object.defineProperty(window, name, original);
    else delete (window as unknown as Record<string, unknown>)[name];
  };
}
