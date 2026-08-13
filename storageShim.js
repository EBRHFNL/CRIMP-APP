/**
 * Demo-only polyfill for the app's `window.storage` interface.
 *
 * The application code (src/App.jsx) calls `window.storage.get/set/delete`
 * to load and save the "Update data" upload. In its original environment
 * that interface is backed by a managed key/value service. Outside that
 * environment, `window.storage` does not exist, so this file provides a
 * small stand-in backed by the browser's localStorage, purely so the demo
 * has working, persistent "Update data" behaviour without needing a
 * backend.
 *
 * Notes:
 * - This is per-browser storage, not shared between users or devices.
 *   For a real multi-user deployment, replace this file with an adapter
 *   that calls a backend API instead (see the technical design document,
 *   section 3.10, for a worked example with Node.js/Express and a small
 *   REST API).
 * - Import this file once, before the app renders (see src/main.jsx).
 */

const STORE_KEY = "hansaflex-crimp-app-demo:kv-store";

function readAll() {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY) || "{}");
  } catch {
    return {};
  }
}

function writeAll(store) {
  localStorage.setItem(STORE_KEY, JSON.stringify(store));
}

function scopedKey(key, shared) {
  return (shared ? "shared:" : "personal:") + key;
}

window.storage = {
  async get(key, shared) {
    const store = readAll();
    const k = scopedKey(key, shared);
    if (!(k in store)) throw new Error("not found");
    return { key, value: store[k], shared };
  },

  async set(key, value, shared) {
    const store = readAll();
    store[scopedKey(key, shared)] = value;
    writeAll(store);
    return { key, value, shared };
  },

  async delete(key, shared) {
    const store = readAll();
    delete store[scopedKey(key, shared)];
    writeAll(store);
    return { key, deleted: true, shared };
  },

  async list(prefix, shared) {
    const store = readAll();
    const scopedPrefix = scopedKey(prefix || "", shared);
    const keys = Object.keys(store)
      .filter((k) => k.startsWith(scopedPrefix))
      .map((k) => k.slice(k.indexOf(":") + 1));
    return { keys };
  },
};
