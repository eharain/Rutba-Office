// The sign-in clients a release build carries (docs/OAUTH.md):
//
//   node build/make-oauth-clients.js
//
// Reads the client file Google's console downloads for a Desktop app
// ({ "installed": { "client_id", "client_secret", ... } }) from
// RUTBA_GOOGLE_CLIENT_FILE, and writes build/oauth/clients.json, which
// electron-builder puts in the resources as oauth/clients.json for
// main/oauth.js to read. The file is made at build time and never committed:
// the id and its "secret" belong to whoever ships the build. With no file
// given, the table is empty and Mail offers an app password, as before.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(APP, 'build', 'oauth', 'clients.json');

/** The id and secret out of Google's download, or a reason it is not one. */
export function googleClient(text) {
  const json = JSON.parse(text);
  const c = json.installed || json.web || {};
  if (!json.installed) throw new Error('not a Desktop app client: the file has no "installed" part');
  if (!/\.apps\.googleusercontent\.com$/.test(c.client_id || '')) throw new Error('no client_id in the file');
  if (!c.client_secret) throw new Error('no client_secret in the file');
  return { clientId: c.client_id, clientSecret: c.client_secret };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const table = {};
  const file = process.env.RUTBA_GOOGLE_CLIENT_FILE;
  if (file) {
    try {
      table.google = googleClient(fs.readFileSync(file, 'utf8'));
    } catch (err) {
      console.error(`make-oauth-clients: ${file}: ${err.message}`);
      process.exit(1);
    }
  }
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, `${JSON.stringify(table)}\n`);
  console.log(table.google
    ? `make-oauth-clients: Google sign-in client ${table.google.clientId.split('-')[0]}-… in the build`
    : 'make-oauth-clients: no RUTBA_GOOGLE_CLIENT_FILE, so this build has no Google sign-in (Mail offers an app password)');
}
