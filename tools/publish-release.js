// Publish a GitHub release from the artefacts in apps/desktop/release.
//
//   node tools/publish-release.js [tag]
//
// The credential comes from git's own store — the same one that pushes the
// commits — so there is no token to paste anywhere and none to leak into a log.
// It is read into memory, used, and never printed.
//
// The upload list matters as much as the installers: `latest.yml` is the feed
// the installed application reads to find a newer version, and the `.blockmap`
// is what lets it download only the parts that changed. A release published
// without them leaves every existing installation on the version it has.

import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const releaseDir = path.join(root, 'apps', 'desktop', 'release');
const OWNER = 'eharain';
const REPO = 'Rutba-Office';

const tag = process.argv[2] || `v${JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version}`;

/** The token git already holds for github.com. */
function credential() {
  const out = execFileSync('git', ['credential', 'fill'], {
    input: 'protocol=https\nhost=github.com\n\n',
    encoding: 'utf8',
    cwd: root,
  });
  const password = /^password=(.*)$/m.exec(out)?.[1];
  if (!password) throw new Error('git has no stored credential for github.com');
  return password;
}

const token = credential();

async function api(url, { method = 'GET', body, headers = {}, raw } = {}) {
  const response = await fetch(url, {
    method,
    headers: {
      authorization: `Bearer ${token}`,
      accept: 'application/vnd.github+json',
      'x-github-api-version': '2022-11-28',
      'user-agent': 'rutba-office-release',
      ...headers,
    },
    body: raw ?? (body ? JSON.stringify(body) : undefined),
    duplex: raw ? 'half' : undefined,
  });
  const text = await response.text();
  let parsed = null;
  try {
    parsed = text ? JSON.parse(text) : null;
  } catch {
    parsed = { raw: text.slice(0, 400) };
  }
  if (!response.ok) {
    const message = parsed?.message || response.statusText;
    const detail = parsed?.errors ? ` — ${JSON.stringify(parsed.errors)}` : '';
    throw new Error(`${method} ${url.replace(/\?.*/, '')} → ${response.status} ${message}${detail}`);
  }
  return parsed;
}

const NOTES = `Seven apps in one download — Mail, Word, Worksheets, Presentation, Pictures,
Image and Video — working with the network switched off.

## What it opens

**Read and write** \`.docx\` \`.xlsx\` \`.pptx\` \`.csv\` \`.tsv\` \`.txt\` \`.md\` \`.html\`
\`.png\` \`.jpg\` \`.webp\` \`.webm\` \`.eml\` \`.mbox\`, and PDF on export.

**Read** \`.doc\` \`.xls\` \`.ppt\` \`.odt\` \`.ods\` \`.odp\` \`.rtf\`, every image and video
format the browser engine decodes, and the mail archives other clients leave
behind: Outlook \`.pst\` and \`.ost\`, Outlook for Mac \`.olm\`, \`.msg\` and \`.emlx\`.

A format that cannot be written is converted on open into one that can, and the
title bar says so — Save As is never a surprise.

## What is different about it

- **It does not damage your files.** The engine rewrites only the parts of a
  document it deliberately edited. Charts, pivot caches, macros, signatures and
  embedded media come back byte-for-byte as they arrived.
- **It works offline.** No account, no sign-in, no telemetry. Automatic update
  checking is off until you turn it on.
- **Your old mail opens.** The \`.pst\` reader was built against a real 906 MB
  \`.ost\` — 467 folders and 12,172 messages — not against fixtures.
- **Mail cannot phone home.** Message bodies render in a frame with no scripts,
  no same-origin access and no remote fetch; remote images load only when asked.

## Downloads

| File | For |
|---|---|
| \`Rutba-Office-1.0.0-win-x64.exe\` | Windows installer |
| \`Rutba-Office-1.0.0-portable.exe\` | Windows, no installation |

macOS and Linux builds are produced by the same configuration and will follow.

**The Windows builds are not code-signed**, so SmartScreen will warn on first
run. The warning is telling you the truth: the publisher is unverified. Check
the file against this page, or build it yourself — the source is here.

## Verified

577 engine tests, 18 application checks and 9 editing checks. The application
checks drive the real windows: open a file, change it, save it, reopen it from
disk, and look at what is actually there. The same run is repeated against the
packaged binary, because that is the only thing that proves what ships works.

Dual-licensed under the GNU AGPL v3.0 and a commercial licence.
[office.rutba.io](https://office.rutba.io)
`;

async function main() {
  if (!fs.existsSync(releaseDir)) throw new Error(`nothing built: ${releaseDir} does not exist`);

  const assets = fs
    .readdirSync(releaseDir)
    .filter((name) => /\.(exe|dmg|zip|AppImage|deb|blockmap)$/.test(name) || /^latest.*\.yml$/.test(name))
    .filter((name) => !name.startsWith('builder-debug'));

  if (!assets.length) throw new Error('no release artefacts found — run npm run dist first');

  console.log(`tag       ${tag}`);
  console.log(`artefacts ${assets.length}`);
  for (const name of assets) {
    console.log(`  ${name.padEnd(44)} ${(fs.statSync(path.join(releaseDir, name)).size / 1048576).toFixed(2)} MB`);
  }

  // Reuse the release if it is already there, so a failed upload can be retried.
  let release = null;
  try {
    release = await api(`https://api.github.com/repos/${OWNER}/${REPO}/releases/tags/${tag}`);
    console.log(`\nrelease already exists (#${release.id}), reusing it`);
  } catch {
    release = await api(`https://api.github.com/repos/${OWNER}/${REPO}/releases`, {
      method: 'POST',
      body: {
        tag_name: tag,
        name: `Rutba Office ${tag.replace(/^v/, '')}`,
        body: NOTES,
        draft: false,
        prerelease: false,
      },
    });
    console.log(`\ncreated release #${release.id}`);
  }

  const existing = new Set((release.assets || []).map((a) => a.name));
  const uploadBase = release.upload_url.replace(/\{.*$/, '');

  for (const name of assets) {
    if (existing.has(name)) {
      console.log(`= ${name} (already uploaded)`);
      continue;
    }
    const file = path.join(releaseDir, name);
    const size = fs.statSync(file).size;
    process.stdout.write(`↑ ${name} (${(size / 1048576).toFixed(1)} MB) … `);
    await api(`${uploadBase}?name=${encodeURIComponent(name)}`, {
      method: 'POST',
      headers: { 'content-type': 'application/octet-stream', 'content-length': String(size) },
      raw: fs.createReadStream(file),
    });
    console.log('done');
  }

  console.log(`\n${release.html_url}`);
}

main().catch((err) => {
  console.error(`\nfailed: ${err.message}`);
  process.exit(1);
});
