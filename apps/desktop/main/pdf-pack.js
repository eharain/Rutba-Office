// PDF's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// The PDF app: its files read once, and new PDFs made from their pages (pdf.js).

import { createPdfService } from './pdf.js';

export function register() {
  return { pdf: createPdfService() };
}
