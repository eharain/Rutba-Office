// Stand-ins for what a computer has one of: its clipboard and its global
// shortcuts. The suite reaches both through adapters (clipboard.js and
// shortcuts.js); a check run is given these, so check processes can run side
// by side, each with a clipboard and shortcuts of its own, and none touches
// the person's. The real adapters are checked on their own, in the gate's
// machine lane, where nothing else runs.
//
// No Electron here: the engine tests check these as they are.

/**
 * A clipboard held in this process, with the system adapter's shape: read()
 * and write() take items with `types` and getType(type) answering a Blob,
 * as Electron's ClipboardItem does; readText, writeText and clear.
 * `Item` builds the items read() answers (Electron's ClipboardItem in the
 * app; a plain stand-in in the tests).
 */
export function fakeClipboard({ Item = null } = {}) {
  let entries = [];
  const makeItem = (parts) => (Item ? new Item(parts) : {
    types: Object.keys(parts),
    getType: async (type) => parts[type],
  });
  const textOf = () => {
    const first = entries.find((e) => e['text/plain']);
    return first ? new TextDecoder().decode(first['text/plain']) : '';
  };
  return {
    kind: 'fake',
    async write(items) {
      entries = await Promise.all((items || []).map(async (item) => {
        const out = {};
        for (const type of item.types) {
          const part = await item.getType(type);
          out[type] = typeof part === 'string' ? new TextEncoder().encode(part) : new Uint8Array(await part.arrayBuffer());
        }
        return out;
      }));
    },
    async read() {
      return entries.map((e) => makeItem(Object.fromEntries(Object.entries(e).map(([type, bytes]) => [type, new Blob([bytes], { type })]))));
    },
    async readText() { return textOf(); },
    async writeText(s) { entries = [{ 'text/plain': new TextEncoder().encode(String(s ?? '')) }]; },
    async clear() { entries = []; },
  };
}

/**
 * Global shortcuts held in this process: register, unregister and
 * isRegistered as Electron's globalShortcut has them, and press(), which a
 * check calls where a person would press the keys anywhere on the computer.
 */
export function fakeShortcuts() {
  const keys = new Map();
  const norm = (a) => String(a).replace(/\s+/g, '').toLowerCase();
  return {
    kind: 'fake',
    register(accelerator, fn) {
      const k = norm(accelerator);
      if (keys.has(k)) return false;
      keys.set(k, fn);
      return true;
    },
    unregister(accelerator) { keys.delete(norm(accelerator)); },
    isRegistered(accelerator) { return keys.has(norm(accelerator)); },
    /** The keys pressed: true when something here was listening for them. */
    press(accelerator) {
      const fn = keys.get(norm(accelerator));
      if (!fn) return false;
      fn();
      return true;
    },
    registered() { return [...keys.keys()]; },
  };
}
