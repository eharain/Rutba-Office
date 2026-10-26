// Which AI tools are offered, from the packs this computer has (docs/APP-PACKS.md 5.1.5).
//
// AI Tools' own tools are about folders and files, and about the mail,
// calendars, address books and databases the person lets them reach; the
// others each make or read one app's kind of file, and are that app's
// pack's. An AI app sees a tool only while its pack is installed and loaded,
// and the apps it reaches too, and none at all without AI Tools. The core writes the packs it loaded into the settings at each
// start ('packs.loaded'), which the tools read for every call, the program
// an AI app starts on this computer as much as the one Rutba Office serves
// to other machines. No list there (a copy from before packs) offers them all.
//
// Standalone on purpose, and in the shell's package so the core and AI
// Tools' pack both reach it by name: the AI tools' program run as Node
// imports this and nothing else of the core.

/** The tools that are another pack's. A tool not named here is AI Tools' own. */
export const TOOL_PACK = {
  create_document: 'word',
  export_pdf: 'word',
  create_workbook: 'sheets',
  create_presentation: 'slides',
  combine_pdf_pages: 'pdf',
  read_text_in_picture: 'ocr',
};

/**
 * The apps an AI Tools tool reaches into (mcp/reach-tools.js): AI Tools'
 * own, offered only while every app it reaches is installed too. An
 * invitation is mailed from the person's own account and kept in a calendar.
 */
export const TOOL_NEEDS = {
  search_mail: ['mail'],
  list_mail_folders: ['mail'],
  list_messages: ['mail'],
  read_message: ['mail'],
  read_attachment: ['mail'],
  create_draft: ['mail'],
  send_message: ['mail'],
  list_events: ['calendar'],
  read_event: ['calendar'],
  create_event: ['calendar'],
  send_invitation: ['calendar', 'mail'],
  answer_invitation: ['calendar', 'mail'],
  search_contacts: ['contacts'],
  read_contact: ['contacts'],
  create_contact: ['contacts'],
  list_tables: ['base'],
  read_records: ['base'],
  add_record: ['base'],
  change_record: ['base'],
};

/** Where the core writes the packs it loaded. */
export const LOADED_KEY = 'packs.loaded';

/** Whether a tool is offered, given the settings (a plain object, as the tools read them). */
export function toolOffered(name, settings) {
  const loaded = settings?.[LOADED_KEY];
  if (!Array.isArray(loaded)) return true;
  if (!loaded.includes('ai')) return false;
  const owner = TOOL_PACK[name];
  if (owner && !loaded.includes(owner)) return false;
  return (TOOL_NEEDS[name] || []).every((id) => loaded.includes(id));
}

/** The tools a pack adds, from the whole list of names. */
export function toolsOf(id, names) {
  return names.filter((n) => (TOOL_PACK[n] || 'ai') === id);
}

/** Why a tool is not offered, in plain words. */
export function toolRefusal(name) {
  return `The tool ${name} is not offered: the app it works with is not installed in Rutba Office on this computer.`;
}
