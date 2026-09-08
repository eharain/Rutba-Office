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

/**
 * The release notes.
 *
 * From `docs/releases/<tag>.md` when there is one, because notes belong with
 * the release they describe rather than inside the script that uploads it —
 * a hard-coded block is one that still says 1.0.0 when 1.4 goes out.
 */
function notesFor(forTag) {
  const file = path.join(root, "docs", "releases", `${forTag}.md`);
  if (fs.existsSync(file)) return fs.readFileSync(file, "utf8");
  throw new Error(
    `No notes for ${forTag}. Write docs/releases/${forTag}.md — a release with no notes is a file nobody knows whether to install.`
  );
}

const NOTES = notesFor(tag);

async function main() {
  if (!fs.existsSync(releaseDir)) throw new Error(`nothing built: ${releaseDir} does not exist`);

  const assets = fs
    .readdirSync(releaseDir)
    .filter((name) => /\.(exe|dmg|zip|AppImage|deb|blockmap)$/.test(name) || /^latest.*\.yml$/.test(name))
    .filter((name) => !name.startsWith('builder-debug'))
    // Only this version's. The build directory keeps every installer ever made
    // on this machine, and a release carrying the previous version's binaries
    // is one where somebody downloads the wrong file — the update feeds are the
    // exception, since they are rewritten in place and always describe the
    // build that has just been made.
    .filter((name) => /^latest.*\.yml$/.test(name) || name.includes(tag.replace(/^v/, '')));

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
        // Beta: every release is a pre-release until the owner says otherwise
        // (RUTBA_RELEASE_FINAL=1 publishes a full release).
        prerelease: process.env.RUTBA_RELEASE_FINAL !== '1',

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
