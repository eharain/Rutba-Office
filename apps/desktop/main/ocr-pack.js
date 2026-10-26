// Text recognition's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// Read on threads of its own: Recognise Text, Copy Text from Picture, Data from Picture (ocr.js).

import { createOcrService } from './ocr.js';

export function register() {
  return { ocr: createOcrService() };
}
