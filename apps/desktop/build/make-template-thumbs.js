// The template gallery's pictures, drawn by the application itself.
//
//   node build/make-template-thumbs.js
//
// Opens each template of packages/templates in the app that makes it (a
// profile of its own, off the desktop, as the checks run) and photographs
// its first page, its grid or its first slide into packages/templates/thumbs.
// Run after tools/make-templates.mjs whenever a template changes.
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
process.env.RUTBA_VERIFY_ONLY = 'template-thumbs';
process.env.RUTBA_TEMPLATE_THUMBS = path.resolve(here, '..', '..', '..', 'packages', 'templates', 'thumbs');
await import('./verify-apps-run.js');
