// The AI tools for an AI app on this computer (D-2026-10-16-06, -07), started
// by that app as Rutba Office's own program running as Node:
//
//   ELECTRON_RUN_AS_NODE=1 "Rutba Office.exe" "<app>\main\mcp\node.js" --profile <id>
//
// (an installed copy names resources\mcp\rutba-office-mcp.cjs, a real file
// beside the archive that loads this one from it: command.js). The AI Tools
// window shows that line ready to copy. Run as Node, the program
// opens no window, leaves the copy of Rutba Office that windows use alone,
// listens on no network port, and its standard input and output carry the
// protocol cleanly (the windowed program prints a line of its own on Windows
// and cannot read its standard input). The settings it obeys are the
// application's, read afresh for each call, so the person's switch in Rutba
// Office decides.
//
// Files it reads and makes itself. Mail, calendars, address books and Data
// Files databases it asks the running Rutba Office of its profile for, over
// a local channel (channel.js): with that copy closed, those tools say so,
// and one that needs the person's press says nothing was done. A profile
// with a password is reached only while it is open and unlocked there.

import fs from 'node:fs';
import path from 'node:path';
import { createServer, serve } from './protocol.js';
import { writeConnection, alive } from './access.js';
import { profileDir, profileEntryOf } from './profile-dir.js';
import { createChannelClient } from './channel.js';
import { needsPassword } from '@rutba/office-shell/profiles';
import { makeTools, versionOf, INSTRUCTIONS } from './make-tools.js';

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
  // Mail, calendars, address books and databases: asked of the running Rutba
  // Office of this profile over its channel (channel.js), which holds them.
  const channel = createChannelClient({ folder: dataDir, isAlive: alive });
  const reach = { call: (name, args) => channel.request('call', { name, args }) };
  // A profile with a password: nothing of it, its folders included, while it
  // is not open and unlocked in a running Rutba Office (D-2026-10-26-06).
  const locked = async () => {
    let entry = null;
    try {
      entry = profileEntryOf(dataDir);
    } catch {
      entry = null;
    }
    if (!needsPassword(entry)) return null;
    try {
      if ((await channel.request('hello'))?.unlocked) return null;
    } catch {
      // Not running, or the tools off there: refused below either way.
    }
    return 'This profile has a password, so the AI tools reach nothing of it, its folders included, unless it is open and unlocked in Rutba Office. Open Rutba Office in this profile and enter its password, then ask again.';
  };
  const tools = makeTools({ settings, dataDir, reach, locked });
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
  channel.close();
  process.exit(0);
}

main().catch((err) => {
  process.stderr.write(`${err?.stack || err}\n`);
  process.exit(1);
});
