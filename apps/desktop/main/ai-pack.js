// AI Tools' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// The AI Tools window: the switch, the folders, the log, and what each mail
// account, calendar, address book, database and folder allows the tools
// (ai-tools.js). Those of an app not installed are not there to allow: the
// services are reached through `packs`, and an absent one answers nothing.

import { createAiToolsService } from './ai-tools.js';

/** The services the tools reach, as they stand when asked: those of the apps installed. */
export function servicesFor(packs) {
  return () => ({ mail: packs.service('mail'), calendar: packs.service('calendar'), contacts: packs.service('contacts'), base: packs.service('base'), ocr: packs.service('ocr') });
}

export function register({ stores, broadcast, packs, onRecentRenamed }) {
  // A database renamed from the launcher keeps what it allows the tools, kept by its place too (mcp/reach.js).
  onRecentRenamed?.((from, to) => {
    if (!/\.rbase$/i.test(from)) return;
    const ai = stores.settings.get('ai.tools.databases', {}) || {};
    const was = Object.keys(ai).find((p) => p.toLowerCase() === from.toLowerCase());
    if (was) stores.settings.set('ai.tools.databases', Object.fromEntries(Object.entries(ai).map(([p, v]) => [p === was ? to : p, v])));
  });
  return { ai: createAiToolsService({ stores, broadcast, services: servicesFor(packs) }) };
}
