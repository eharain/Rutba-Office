// PDF's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// The PDF app: its files read once, and new PDFs made from their pages (pdf.js).

import { createPdfService } from './pdf.js';

// The core's in-memory table comes in the context, for a locked PDF opened in memory.
export function register({ memoryFiles } = {}) {
  return { pdf: createPdfService({ memoryFiles }) };
}
