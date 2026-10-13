// The template gallery's list: what each template is, which app makes it,
// the categories it is found under, and the file it is.
//
// Plain data, so the windows, the main process, the generator and the tests
// all read the same list. The words are English; the windows show them
// through t() (renderer/template-gallery.js lists every one of them, so the
// message catalogue carries them in Urdu, Arabic and Hindi too). What is
// inside a template, its own headings and figures, is English for now.

/** The categories, in the order the gallery lists them. */
export const CATEGORIES = [
  { id: 'business', name: 'Business' },
  { id: 'finance', name: 'Finance' },
  { id: 'planning', name: 'Planning' },
  { id: 'education', name: 'Education' },
  { id: 'career', name: 'Career' },
  { id: 'household', name: 'Household' },
];

/**
 * Every template. `file` is the template's name under files/, `thumb` its
 * picture under thumbs/. `legacy` names the quick start on the launcher
 * (registry NEW_DOCUMENTS) this template now stands behind.
 */
export const TEMPLATES = [
  { id: 'word-letter', app: 'word', ext: 'docx', name: 'Letter', description: 'A business letter: your details, the date, the address, a subject line and a signature block', categories: ['business', 'career'], legacy: 'letter' },
  { id: 'word-cv', app: 'word', ext: 'docx', name: 'CV', description: 'A one-page CV: a profile, experience, education and skills', categories: ['career'] },
  { id: 'word-report', app: 'word', ext: 'docx', name: 'Report', description: 'A report with a cover page, a table of contents and numbered sections', categories: ['business', 'education'], legacy: 'report' },
  { id: 'word-invoice', app: 'word', ext: 'docx', name: 'Invoice', description: 'An invoice to print or send: who it is from and to, the items, the total and how to pay', categories: ['business', 'finance'] },
  { id: 'word-minutes', app: 'word', ext: 'docx', name: 'Meeting minutes', description: 'Minutes of a meeting: who came, what was discussed and decided, and who does what next', categories: ['business', 'planning'] },
  { id: 'word-newsletter', app: 'word', ext: 'docx', name: 'Newsletter', description: 'A newsletter with a masthead, two columns of stories and a dates panel', categories: ['business', 'household'] },
  { id: 'sheet-budget', app: 'sheets', ext: 'xlsx', name: 'Monthly budget', description: 'Income and spending, planned against actual, with charts of where the money goes', categories: ['finance', 'household'], legacy: 'budget' },
  { id: 'sheet-invoice', app: 'sheets', ext: 'xlsx', name: 'Invoice', description: 'An invoice that adds itself up: quantities, prices, tax and the total due', categories: ['business', 'finance'], legacy: 'invoice' },
  { id: 'sheet-timesheet', app: 'sheets', ext: 'xlsx', name: 'Timesheet', description: 'A week of hours: start, finish and breaks, with the hours and pay worked out', categories: ['business', 'planning'] },
  { id: 'sheet-loan', app: 'sheets', ext: 'xlsx', name: 'Loan calculator', description: 'The monthly payment, the total interest and a year-by-year schedule for a loan', categories: ['finance', 'household'] },
  { id: 'sheet-todo', app: 'sheets', ext: 'xlsx', name: 'To-do list', description: 'A table of tasks with owners, due dates, priorities and progress', categories: ['planning', 'household'] },
  { id: 'sheet-ledger', app: 'sheets', ext: 'xlsx', name: 'Accounts ledger', description: 'A simple cash book: money in, money out and the running balance', categories: ['finance', 'business'] },
  { id: 'deck-pitch', app: 'slides', ext: 'pptx', name: 'Pitch deck', description: 'The story of a business in eight slides: problem, product, market, traction, team and the ask', categories: ['business'], legacy: 'pitch' },
  { id: 'deck-status', app: 'slides', ext: 'pptx', name: 'Project status', description: 'Where a project stands: progress, milestones, risks and next steps', categories: ['business', 'planning'] },
  { id: 'deck-album', app: 'slides', ext: 'pptx', name: 'Photo album', description: 'Pictures with captions, one large and in pairs, on a dark theme', categories: ['household'] },
  { id: 'deck-training', app: 'slides', ext: 'pptx', name: 'Training deck', description: 'A course session: objectives, the agenda, the lessons, a quick check and a summary', categories: ['education', 'business'] },
].map((t) => ({ ...t, file: `${t.id}.${t.ext}`, thumb: `${t.id}.png` }));

/** A template by its id, or null. */
export function templateById(id) {
  return TEMPLATES.find((t) => t.id === id) || null;
}

/** The gallery template behind a launcher quick start ('letter', 'budget' …) in an app, or null. */
export function templateForLegacy(app, legacy) {
  return TEMPLATES.find((t) => t.legacy === legacy && t.app === app) || null;
}

/** The templates an app makes. */
export function templatesOf(app) {
  return TEMPLATES.filter((t) => t.app === app);
}
