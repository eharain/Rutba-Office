// The template gallery's files, made by the suite's own engines.
//
//   node tools/make-templates.mjs        writes packages/templates/files/*
//
// Each template's specification is in packages/templates/src (word.js,
// sheets.js, decks.js). The thumbnails are drawn by the application itself:
// `node build/make-template-thumbs.js` in apps/desktop, after this.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeAll } from '@rutba/templates/make';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(ROOT, 'packages', 'templates', 'files');
fs.mkdirSync(out, { recursive: true });
for (const { entry, bytes } of makeAll()) {
  fs.writeFileSync(path.join(out, entry.file), bytes);
  console.log(`${entry.file}  ${bytes.length} bytes`);
}
