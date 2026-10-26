// AI Tools' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// The AI Tools window: the switch, the folders, the log, and what each mail
// account, calendar, address book, database and folder allows the tools
// (ai-tools.js). Those of an app not installed are not there to allow: the
// services are reached through `packs`, and an absent one answers nothing.
//
// The program an AI app starts on this computer is AI Tools' too (`program`,
// mcp/program.js): the core's start file for it (mcp/node.js) runs it from
// here. That program runs as Node, without Electron, so the window's service
// (ai-tools.js, which needs Electron) is loaded only when it is registered.

/** The services the tools reach, as they stand when asked: those of the apps installed. */
export function servicesFor(packs) {
  return () => ({ mail: packs.service('mail'), calendar: packs.service('calendar'), contacts: packs.service('contacts'), base: packs.service('base'), ocr: packs.service('ocr') });
}

export async function register({ stores, broadcast, packs, onRecentRenamed, unicodeFont }) {
  const { createAiToolsService } = await import('./ai-tools.js');
  // A database renamed from the launcher keeps what it allows the tools, kept by its place too (mcp/reach.js).
  onRecentRenamed?.((from, to) => {
    if (!/\.rbase$/i.test(from)) return;
    const ai = stores.settings.get('ai.tools.databases', {}) || {};
    const was = Object.keys(ai).find((p) => p.toLowerCase() === from.toLowerCase());
    if (was) stores.settings.set('ai.tools.databases', Object.fromEntries(Object.entries(ai).map(([p, v]) => [p === was ? to : p, v])));
  });
  return { ai: createAiToolsService({ stores, broadcast, services: servicesFor(packs), unicodeFont }) };
}

/** The tools for an AI app on this computer, served over standard input and output (mcp/program.js). */
export async function program(core) {
  return (await import('./mcp/program.js')).program(core);
}
