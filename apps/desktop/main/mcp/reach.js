// What the AI tools may reach, thing by thing (owner, 2026-10-26,
// D-2026-10-26-04, -05, -06: "the mcp tool should also let the user choose
// what to allow from apps, what emails accounts to allow to all etc.", "so
// the settings should be email accounts", "same for calendars and contacts
// and across the board").
//
// Every thing the tools could reach carries its own setting, kept with it:
//   a mail account       `ai` on its record in settings (mail.accounts)
//   an archive opened    `ai` on its entry in settings (mail.archives)
//   a calendar           `ai` on the calendar in calendar.json
//   an address book      `ai` on the book in contacts.json (`ai` of the file
//                        itself for the cards kept on this computer only)
//   a Data Files database  settings ai.tools.databases, by its path, with the
//                        role the tools work as when it is protected: kept on
//                        this computer, never in the file, as what it may
//                        read is (base-trust.js)
//   a folder of files    settings ai.tools.folders: a path alone is a folder
//                        given before 2.19, which keeps what it allowed then
//                        (read, and new files); { path, step } one set since
// The same four steps everywhere, and Off until the person chooses:
//   off      nothing
//   read     read and search
//   prepare  read, and make things kept on this computer for the person to
//            look at: drafts, new events and contacts, new files and records
//   act      read, prepare, and do what leaves this computer or tells another
//            person (a message sent, an invitation or its answer, a contact
//            written to a server, a change written to a database engine),
//            each shown in Rutba Office first and done only when the person
//            presses it (confirm.js)
// Deleting or moving anything is never in reach, whatever the step.
//
// Nothing here needs Electron or reads a file: the AI Tools service
// (ai-tools.js) hands in what it reads, and the tools ask these rules.

export const STEPS = ['off', 'read', 'prepare', 'act'];
export const KINDS = ['mail', 'archive', 'calendar', 'book', 'database', 'folder'];
export const FOLDERS_KEY = 'ai.tools.folders';
export const DATABASES_KEY = 'ai.tools.databases';
/** The address book of the cards kept on this computer only, which has no record of its own. */
export const LOCAL_BOOK = 'local';

/** What each step is called, in the tools' answers (the windows say it in their own words). */
export const STEP_WORDS = { off: 'Off', read: 'Read', prepare: 'Read and prepare', act: 'Read, prepare and act' };

/**
 * The steps a thing may be given: an archive opened in place is read only;
 * a mailbox brought in from a file has no server to send through, and a
 * folder nothing to send, so neither goes past prepare.
 */
export function stepsFor(kind, thing = null) {
  if (kind === 'archive') return ['off', 'read'];
  if (kind === 'folder') return ['off', 'read', 'prepare'];
  if (kind === 'mail' && thing?.local) return ['off', 'read', 'prepare'];
  return STEPS;
}

/** A step as kept, made one this thing may have: anything else is Off. */
export function stepOf(kind, value, thing = null) {
  return stepsFor(kind, thing).includes(value) ? value : 'off';
}

/** Whether `step` allows what `want` needs. */
export const allows = (step, want) => STEPS.indexOf(step) >= STEPS.indexOf(want) && want !== 'off';

/** Where each kind's setting is, as a refusal says it. */
const WHERE = {
  mail: 'in Mail under Accounts, or in AI Tools',
  archive: 'in AI Tools',
  calendar: 'in Calendar under Accounts, or in AI Tools',
  book: 'in Contacts under Accounts, or in AI Tools',
  database: 'in Data Files (the database\'s AI tools, in its side panel), or in AI Tools',
  folder: 'in AI Tools, under Folders',
};
const NOUN = { mail: 'the mail account', archive: 'the archive', calendar: 'the calendar', book: 'the address book', database: 'the database', folder: 'the folder' };
const NEED = {
  read: 'Read',
  prepare: 'Read and prepare',
  act: 'Read, prepare and act',
};

/** Where a kind's setting is changed, in a sentence's words. */
export const whereOf = (kind) => WHERE[kind] || 'in AI Tools';

/**
 * The sentence a refusal says: the thing by name, what it is set to, what the
 * call needs, and where the setting is.
 */
export function refusalFor(kind, name, step, want) {
  const thing = `${NOUN[kind] || 'that'} ${name}`;
  if (!stepsFor(kind).includes(want)) {
    if (kind === 'archive') return `The AI tools only read ${thing}: an archive opened in place is never changed.`;
    return `The AI tools cannot do that with ${thing}.`;
  }
  return `The AI tools may not do that with ${thing}: it is set to ${STEP_WORDS[step] || 'Off'}, and this needs ${NEED[want]}. Its setting is ${whereOf(kind)}.`;
}

/** A refusal as an error the tools turn into their answer. */
export function refuse(kind, name, step, want) {
  // What it was refused for goes in the log, as a call that worked would: the thing and the step it needed.
  return Object.assign(new Error(refusalFor(kind, name, step, want)), { refusal: true, log: { kind, name, step: want } });
}

/* ── folders ─────────────────────────────────────────────────────────── */

/**
 * The folders as kept, each with its step: a path alone (given before 2.19)
 * keeps what it allowed then, read and new files; Off folders are listed
 * too, so the window shows them, and reach nothing.
 */
export function folderEntries(list) {
  const out = [];
  const seen = new Set();
  for (const f of Array.isArray(list) ? list : []) {
    const p = typeof f === 'string' ? f : typeof f?.path === 'string' ? f.path : null;
    if (!p) continue;
    const key = process.platform === 'win32' ? p.toLowerCase() : p;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ path: p, step: typeof f === 'string' ? 'prepare' : stepOf('folder', f.step) });
  }
  return out;
}

/** The list kept again with one folder's step changed (or the folder added). */
export function withFolderStep(list, folder, step, same = (a, b) => a.toLowerCase() === b.toLowerCase()) {
  const entries = folderEntries(list);
  const at = entries.findIndex((e) => same(e.path, folder));
  const next = { path: folder, step: stepOf('folder', step) };
  if (at >= 0) entries[at] = next;
  else entries.push(next);
  // A folder set to what a path alone means is kept as a path alone, as before 2.19.
  return entries.map((e) => (e.step === 'prepare' ? e.path : e));
}

/* ── databases ───────────────────────────────────────────────────────── */

/** One database's setting as kept: { step, role }. */
export function databaseSetting(map, file, same = (a, b) => a.toLowerCase() === b.toLowerCase()) {
  for (const [p, v] of Object.entries(map && typeof map === 'object' ? map : {})) {
    if (same(p, file)) return { step: stepOf('database', v?.step), role: typeof v?.role === 'string' ? v.role : null };
  }
  return { step: 'off', role: null };
}
