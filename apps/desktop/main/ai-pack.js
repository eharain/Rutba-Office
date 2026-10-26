// AI Tools' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// The AI Tools window: the switch, the folders, the log (ai-tools.js).

import { createAiToolsService } from './ai-tools.js';

export function register({ stores }) {
  return { ai: createAiToolsService({ stores }) };
}
