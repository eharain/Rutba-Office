// The gallery's files on disk, for the main process: where each template and
// its thumbnail are kept. The windows read the list from ./catalogue.js and
// ask the main process for the rest.

import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { TEMPLATES, CATEGORIES, templateById, templateForLegacy, templatesOf } from './catalogue.js';

export { TEMPLATES, CATEGORIES, templateById, templateForLegacy, templatesOf };

/** The folder the templates are kept in. */
export const FILES_DIR = fileURLToPath(new URL('../files/', import.meta.url));
/** The folder their thumbnails are kept in. */
export const THUMBS_DIR = fileURLToPath(new URL('../thumbs/', import.meta.url));

/** A template's file, as a path. */
export function templatePath(entry) {
  return FILES_DIR + entry.file;
}

/** A template's thumbnail, as a path. */
export function thumbPath(entry) {
  return THUMBS_DIR + entry.thumb;
}

/** A template's bytes; throws a sentence when the file is not there. */
export function readTemplate(id) {
  const entry = templateById(id);
  if (!entry) throw new Error(`There is no template "${id}".`);
  return fs.readFileSync(templatePath(entry));
}
