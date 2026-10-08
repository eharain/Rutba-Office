// Put the release notes on GitHub again, from docs/releases, for releases
// already published. Only each release's text changes; its name, tag,
// pre-release flag and files are left as they are.
//
//   node tools/update-release-notes.js --dry-run --all   say what would change
//   node tools/update-release-notes.js --all             every release with a notes file
//   node tools/update-release-notes.js v1.29.6 v1.29.7   just these
//
// The body goes up the way publish-release.js sends it: each paragraph and
// list item on one line (release-notes.js), because the release page breaks
// a line wherever the body does. A release with no notes file (1.0.0 was
// written in its tag) is left alone. The credential is git's own, as in
// publish-release.js, read into memory and never printed.

import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { unwrapNotes } from './release-notes.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OWNER = 'eharain';
const REPO = 'Rutba-Office';

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const all = args.includes('--all');
const named = args.filter((a) => /^v\d+\.\d+\.\d+$/.test(a));
if (!all && !named.length) {
  console.error('Name the releases (v1.29.7 …) or pass --all; add --dry-run to see what would change.');
  process.exit(2);
}

function credential() {
  try {
    const out = execFileSync('git', ['credential', 'fill'], {
      input: 'protocol=https\nhost=github.com\n\n',
      encoding: 'utf8',
      cwd: root,
      stdio: ['pipe', 'pipe', 'ignore'],
    });
    const password = /^password=(.*)$/m.exec(out)?.[1];
    if (password) return password;
  } catch {
    /* no stored credential; the environment may carry one */
  }
  const fromEnv = process.env.GH_TOKEN || process.env.GITHUB_TOKEN;
  if (fromEnv) return fromEnv;
  throw new Error('git has no stored credential for github.com, and neither GH_TOKEN nor GITHUB_TOKEN is set');
}

const token = dryRun ? null : credential();

async function api(url, { method = 'GET', body } = {}) {
  const headers = {
    accept: 'application/vnd.github+json',
    'x-github-api-version': '2022-11-28',
    'user-agent': 'rutba-office-release',
  };
  if (token) headers.authorization = `Bearer ${token}`;
  if (body) headers['content-type'] = 'application/json';
  const response = await fetch(url, { method, headers, body: body ? JSON.stringify(body) : undefined });
  if (!response.ok) throw new Error(`${method} ${url}: ${response.status} ${(await response.text()).slice(0, 200)}`);
  return response.json();
}

async function releases() {
  const list = [];
  for (let page = 1; ; page++) {
    const batch = await api(`https://api.github.com/repos/${OWNER}/${REPO}/releases?per_page=100&page=${page}`);
    list.push(...batch);
    if (batch.length < 100) return list;
  }
}

const published = await releases();
const wanted = all ? published : published.filter((r) => named.includes(r.tag_name));
let changed = 0;
for (const release of wanted) {
  const file = path.join(root, 'docs', 'releases', `${release.tag_name}.md`);
  if (!fs.existsSync(file)) {
    console.log(`${release.tag_name.padEnd(9)} no notes file, left as it is`);
    continue;
  }
  const body = unwrapNotes(fs.readFileSync(file, 'utf8')).trim();
  if ((release.body || '').replace(/\r\n/g, '\n').trim() === body) {
    console.log(`${release.tag_name.padEnd(9)} already current`);
    continue;
  }
  changed++;
  if (dryRun) {
    console.log(`${release.tag_name.padEnd(9)} would change (${(release.body || '').length} → ${body.length} characters)`);
    continue;
  }
  await api(`https://api.github.com/repos/${OWNER}/${REPO}/releases/${release.id}`, { method: 'PATCH', body: { body } });
  console.log(`${release.tag_name.padEnd(9)} updated`);
}
console.log(`${dryRun ? 'would update' : 'updated'} ${changed} of ${wanted.length}`);
