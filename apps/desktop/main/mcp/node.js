// The AI tools for an AI app on this computer (D-2026-10-16-06, -07), started
// by that app as Rutba Office's own program running as Node:
//
//   ELECTRON_RUN_AS_NODE=1 "Rutba Office.exe" "<app>\main\mcp\node.js"
//
// The AI Tools window shows that line ready to copy. Run as Node, the program
// opens no window, leaves the copy of Rutba Office that windows use alone,
// listens on no network port, and its standard input and output carry the
// protocol cleanly (the windowed program prints a line of its own on Windows
// and cannot read its standard input). The settings it obeys are the
// application's, read afresh for each call, so the person's switch in Rutba
// Office decides.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { isPng } from '@rutba/ocr/png';
import { createServer, serve } from './protocol.js';
import { createTools } from './tools.js';
import { writeConnection } from './access.js';
import { unicodeFont } from '../system-fonts.js';
import { createOcrService } from '../ocr.js';

const INSTRUCTIONS = 'Rutba Office makes and reads documents (.docx), workbooks (.xlsx), presentations (.pptx) and PDFs on this computer, in the folders its user has given it, and never overwrites or deletes a file. The tools are off unless the user has turned them on in Rutba Office (app menu, AI Tools); a refusal says so.';

/** Rutba Office's profile folder, where its settings are: as Electron names it for the app. */
export function profileDir() {
  if (process.env.RUTBA_OFFICE_PROFILE) return process.env.RUTBA_OFFICE_PROFILE;
  if (process.platform === 'win32') return path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Rutba Office');
  if (process.platform === 'darwin') return path.join(os.homedir(), 'Library', 'Application Support', 'Rutba Office');
  return path.join(process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config'), 'Rutba Office');
}

/** The version this copy is, from the app's package. */
function versionOf() {
  try {
    return JSON.parse(fs.readFileSync(new URL('../../package.json', import.meta.url), 'utf8')).version || '0.0.0';
  } catch {
    return '0.0.0';
  }
}

export async function main() {
  // Standard output carries the protocol and nothing else.
  const toErr = (...a) => process.stderr.write(`${a.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' ')}\n`);
  console.log = toErr;
  console.info = toErr;
  console.warn = toErr;
  console.debug = toErr;

  const dataDir = profileDir();
  const settingsFile = path.join(dataDir, 'settings.json');
  const settings = () => {
    try {
      return JSON.parse(fs.readFileSync(settingsFile, 'utf8'));
    } catch {
      return {};
    }
  };
  let ocr = null;
  const tools = createTools({
    settings,
    dataDir,
    renderPdf,
    openDocx,
    unicodeFont,
    ocr: { recognize: (args) => (ocr ||= createOcrService()).recognize(args) },
    pictureToPng: async (file) => {
      const bytes = new Uint8Array(fs.readFileSync(file));
      if (!isPng(bytes)) throw new Error(`${path.basename(file)} is not a PNG picture. The AI tools read the words in PNG pictures; save it as a PNG first.`);
      return bytes;
    },
  });
  // Who is connected, for the AI Tools window and the badge in every window's
  // title bar: what the AI app says it is, from where, since when, its last call.
  const pid = process.pid;
  const since = new Date().toISOString();
  process.on('exit', () => writeConnection(dataDir, pid, null));
  const server = createServer({
    name: 'rutba-office',
    version: versionOf(),
    tools,
    instructions: INSTRUCTIONS,
    onInitialize: ({ client }) => writeConnection(dataDir, pid, { client: String(client?.name || 'An AI app').slice(0, 80), clientVersion: String(client?.version || '').slice(0, 40), where: 'this computer', since, last: null }),
    onCall: (tool) => writeConnection(dataDir, pid, { last: new Date().toISOString(), lastTool: String(tool || '').slice(0, 60) }),
  });
  await serve(server, process.stdin, process.stdout);
  process.exit(0);
}

main().catch((err) => {
  process.stderr.write(`${err?.stack || err}\n`);
  process.exit(1);
});
