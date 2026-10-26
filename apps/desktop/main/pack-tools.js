// Which AI tools are offered, from the packs this computer has (docs/APP-PACKS.md 5.1.5).
//
// AI Tools' own tools are about folders and files; the others each make or
// read one app's kind of file, and are that app's pack's: an AI app sees a
// tool only while its pack is installed and loaded, and none at all without
// AI Tools. The core writes the packs it loaded into the settings at each
// start ('packs.loaded'), which the tools read for every call, the program
// an AI app starts on this computer as much as the one Rutba Office serves
// to other machines. No list there (a copy from before packs) offers them all.
//
// Standalone on purpose: the AI tools' program run as Node imports this and
// nothing else of the core.

/** The tools that are another pack's. A tool not named here is AI Tools' own. */
export const TOOL_PACK = {
  create_document: 'word',
  export_pdf: 'word',
  create_workbook: 'sheets',
  create_presentation: 'slides',
  combine_pdf_pages: 'pdf',
  read_text_in_picture: 'ocr',
};

/** Where the core writes the packs it loaded. */
export const LOADED_KEY = 'packs.loaded';

/** Whether a tool is offered, given the settings (a plain object, as the tools read them). */
export function toolOffered(name, settings) {
  const loaded = settings?.[LOADED_KEY];
  if (!Array.isArray(loaded)) return true;
  if (!loaded.includes('ai')) return false;
  const owner = TOOL_PACK[name];
  return !owner || loaded.includes(owner);
}

/** The tools a pack adds, from the whole list of names. */
export function toolsOf(id, names) {
  return names.filter((n) => (TOOL_PACK[n] || 'ai') === id);
}

/** Why a tool is not offered, in plain words. */
export function toolRefusal(name) {
  return `The tool ${name} is not offered: the app it works with is not installed in Rutba Office on this computer.`;
}
