// The AI tools for an AI app on this computer (D-2026-10-16-06, -07), started
// by that app as Rutba Office's own program running as Node:
//
//   ELECTRON_RUN_AS_NODE=1 "Rutba Office.exe" "<app>\main\mcp\node.js" --profile <id>
//
// (an installed copy names resources\mcp\rutba-office-mcp.cjs, a real file
// beside the archive that loads this one from it: command.js). The AI Tools
// window shows that line ready to copy.
//
// This file is the core's: it finds the packs this computer has, checked as a
// start of Rutba Office checks them, and runs AI Tools' program from AI
// Tools' pack (program.js), handing it this copy's version, text recognition
// when that pack is installed, and the system font a PDF is written in. With
// AI Tools not installed it says so and ends.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { planPacks, readPackKeys, packModule, packService } from '../app-packs.js';
import { unicodeFont } from '../system-fonts.js';
import { installPackResolution } from '@rutba/office-shell/electron/pack-resolve';

/** The version this copy is, from the app's package. */
function versionOf(app) {
  try {
    return JSON.parse(fs.readFileSync(path.join(app, 'package.json'), 'utf8')).version || '0.0.0';
  } catch {
    return '0.0.0';
  }
}

/** The packs this computer has: the folder beside the archive, or a check run's; none in a copy run from the source. */
function plan() {
  const app = fileURLToPath(new URL('../../', import.meta.url));
  const packaged = /\.asar[\\/]?$/.test(app);
  const named = process.env.RUTBA_PACKS_DIR ? path.resolve(process.env.RUTBA_PACKS_DIR) : null;
  const beside = packaged ? path.join(app, '..', 'packs') : null;
  const dir = named || (beside && fs.existsSync(beside) ? beside : null);
  const keys = readPackKeys(path.join(app, 'build', 'out', 'pack-keys.json'));
  if (!packaged && process.env.RUTBA_PACK_PUBLIC_KEY) keys.push(fs.readFileSync(process.env.RUTBA_PACK_PUBLIC_KEY, 'utf8'));
  const version = versionOf(app);
  if (dir) installPackResolution({ packsDir: dir, coreFile: fileURLToPath(import.meta.url) });
  return { version, packs: planPacks({ dir, keys, version }) };
}

export async function main() {
  const { version, packs } = plan();
  const ai = await packModule(packs, 'ai');
  if (!ai?.program) {
    process.stderr.write('AI Tools is not installed in Rutba Office on this computer. Add it with Add or remove apps.\n');
    process.exit(1);
  }
  let ocr = null;
  await ai.program({
    version,
    // Text recognition, for the tool that reads the words in a picture: from its pack, when it is installed.
    ocr: () => (ocr ||= packService(packs, 'ocr').then((s) => s?.ocr || null)),
    unicodeFont,
  });
  process.exit(0);
}

main().catch((err) => {
  process.stderr.write(`${err?.stack || err}\n`);
  process.exit(1);
});
