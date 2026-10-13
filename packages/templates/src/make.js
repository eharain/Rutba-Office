// Making the gallery's files: each template's bytes, from its specification,
// through the suite's own engines. tools/make-templates.mjs writes them into
// files/; tests/templates.test.js makes them again and holds the files kept
// to what this makes, so a template is never a file nobody can remake.

import { TEMPLATES } from './catalogue.js';
import { DOCUMENTS } from './word.js';
import { WORKBOOKS } from './sheets.js';
import { DECKS } from './decks.js';

const MAKERS = { ...DOCUMENTS, ...WORKBOOKS, ...DECKS };

/** One template's bytes, made now. */
export function makeTemplate(id) {
  const make = MAKERS[id];
  if (!make) throw new Error(`no maker for the template "${id}"`);
  return make();
}

/** Every template, made now: [{ entry, bytes }]. */
export function makeAll() {
  return TEMPLATES.map((entry) => ({ entry, bytes: makeTemplate(entry.id) }));
}
