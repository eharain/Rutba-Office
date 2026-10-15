// Publish a GitHub release from the artefacts in apps/desktop/release.
//
//   node tools/publish-release.js [tag]
//
// The credential comes from git's own store (the same one that pushes the
// commits), so there is no token to paste anywhere and none to leak into a log.
// It is read into memory, used, and never printed.
//
// The upload list matters as much as the installers: `latest.yml` is the feed
// the installed application reads to find a newer version, and the `.blockmap`
// is what lets it download only the parts that changed. A release published
// without them leaves every existing installation on the version it has.
//
// So the release is made as a DRAFT, and published only when its files are
// there and agree with each other (1.44.0's lesson: its release went up with
// no installers, and every copy of 1.43.0 found a release with no update file
// in it). In order:
//
//   1. the draft: found again if an earlier run left one (a failed upload is
//      retried by running this again), else made. A draft is seen by nobody:
//      the updater's feed and the releases page leave it out, and GitHub makes
//      no tag for it, so a release run from the Actions tab tags the commit
//      only at step 4;
//   2. every file uploaded, a file a failed run left half-uploaded or at the
//      wrong size replaced;
//   3. the check: the release's files listed back from GitHub, each there,
//      uploaded and the size it is here; and each update feed (latest*.yml)
//      read back from the release itself and held to the files: its version
//      the tag's, every file it names uploaded, at the size it says, with the
//      SHA-512 of the file that was built, and each installer's blockmap
//      beside it. Anything wrong stops here, and the draft stays a draft;
//   4. published, as a pre-release unless RUTBA_RELEASE_FINAL=1.
//
// A release already published (made before this, or by hand) is not taken
// back to a draft: what is missing is uploaded to it and it is checked the same way.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { unwrapNotes } from './release-notes.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OWNER = 'eharain';
const REPO = 'Rutba-Office';

/** The files of this version in a build directory, in the order they go up. */
export function releaseAssets(names, tag) {
  const version = tag.replace(/^v/, '');
  return names
    .filter((name) => /\.(exe|dmg|zip|AppImage|deb|blockmap)$/.test(name) || /^latest.*\.yml$/.test(name))
    .filter((name) => !name.startsWith('builder-debug'))
    // Only this version's. The build directory keeps every installer ever made
    // on this machine, and a release carrying the previous version's binaries
    // is one where somebody downloads the wrong file: the update feeds are the
    // exception, since they are rewritten in place and always describe the
    // build that has just been made.
    .filter((name) => /^latest.*\.yml$/.test(name) || name.includes(version))
    // The feeds last: an updater that finds a feed finds its files already there.
    .sort((a, b) => Number(/^latest.*\.yml$/.test(a)) - Number(/^latest.*\.yml$/.test(b)) || a.localeCompare(b));
}

/**
 * electron-builder's update feed (latest.yml, latest-mac.yml,
 * latest-linux.yml): its scalars, and `files` as a list of maps. Only the
 * shape electron-builder writes, which is all this has to read.
 */
export function parseFeed(text) {
  const feed = { files: [] };
  let inFiles = false;
  let current = null;
  const scalar = (v) => {
    const s = v.trim();
    if (/^'.*'$/.test(s)) return s.slice(1, -1).replace(/''/g, "'");
    if (/^".*"$/.test(s)) return s.slice(1, -1);
    if (/^-?\d+$/.test(s)) return Number(s);
    if (s === 'true' || s === 'false') return s === 'true';
    return s;
  };
  for (const line of String(text).replace(/\r\n/g, '\n').split('\n')) {
    if (!line.trim() || line.trim().startsWith('#')) continue;
    const top = /^([A-Za-z][\w-]*):\s*(.*)$/.exec(line);
    if (top) {
      inFiles = top[1] === 'files' && top[2] === '';
      current = null;
      if (!inFiles) feed[top[1]] = scalar(top[2]);
      continue;
    }
    const item = /^\s+-\s+([A-Za-z][\w-]*):\s*(.*)$/.exec(line);
    if (inFiles && item) {
      current = { [item[1]]: scalar(item[2]) };
      feed.files.push(current);
      continue;
    }
    const field = /^\s+([A-Za-z][\w-]*):\s*(.*)$/.exec(line);
    if (inFiles && current && field) current[field[1]] = scalar(field[2]);
  }
  return feed;
}

/** A file's SHA-512 as an update feed writes it: base64. */
export const sha512Of = (bytes) => crypto.createHash('sha512').update(bytes).digest('base64');

/**
 * What is wrong between an update feed and the release's files, as
 * sentences; empty when they agree. `uploaded` is name to size as GitHub
 * lists them; `local` is name to { size, sha512 } of the files built here.
 */
export function checkFeed({ name, feed, tag, uploaded, local }) {
  const problems = [];
  const version = tag.replace(/^v/, '');
  if (String(feed.version) !== version) problems.push(`${name} is for version ${feed.version}, not ${version}`);
  const files = feed.files?.length ? feed.files : feed.path ? [{ url: feed.path, sha512: feed.sha512, size: feed.size }] : [];
  if (!files.length) problems.push(`${name} names no file to update with`);
  for (const f of files) {
    const fileName = decodeURIComponent(String(f.url || '')).split('/').pop();
    if (!uploaded.has(fileName)) {
      problems.push(`${name} names ${fileName}, which is not in the release`);
      continue;
    }
    const here = local.get(fileName);
    if (f.size != null && uploaded.get(fileName) !== f.size) problems.push(`${name} says ${fileName} is ${f.size} bytes; the release has ${uploaded.get(fileName)}`);
    if (here && f.sha512 !== here.sha512) problems.push(`${name}'s SHA-512 for ${fileName} is not the SHA-512 of the file that was built`);
    // electron-updater downloads a Windows installer by its blockmap's parts.
    if (/\.exe$/.test(fileName) && !uploaded.has(`${fileName}.blockmap`)) {
      problems.push(`${fileName}.blockmap is not in the release, so an update would download the whole installer`);
    }
  }
  if (feed.path && feed.sha512 && files[0] && feed.sha512 !== files[0].sha512) problems.push(`${name}'s path and its first file disagree`);
  return problems;
}

/**
 * The release, made and published as the header says. `client` is GitHub
 * (githubClient below), or the engine tests' stand-in for it.
 */
export async function publishRelease({ client, dir, tag, notes, commit = null, prerelease = true, log = console.log }) {
  if (!fs.existsSync(dir)) throw new Error(`nothing built: ${dir} does not exist`);
  const assets = releaseAssets(fs.readdirSync(dir), tag);
  if (!assets.length) throw new Error('no release artefacts found: run npm run dist first');
  if (!assets.some((a) => /^latest.*\.yml$/.test(a))) throw new Error('there is no update feed (latest.yml) beside the installers, so an installed copy would never find this release');

  log(`tag       ${tag}`);
  log(`artefacts ${assets.length}`);
  const local = new Map();
  for (const name of assets) {
    const bytes = fs.readFileSync(path.join(dir, name));
    local.set(name, { size: bytes.length, sha512: sha512Of(bytes) });
    log(`  ${name.padEnd(44)} ${(bytes.length / 1048576).toFixed(2)} MB`);
  }

  // 1. The draft: one an earlier run left, or a new one.
  let release = (await client.listReleases()).find((r) => r.tag_name === tag) || null;
  if (release) log(`\nrelease already exists (#${release.id}, ${release.draft ? 'a draft' : 'published'}), reusing it`);
  else {
    release = await client.createRelease({
      tag_name: tag,
      // The tag, if it is not pushed yet, is made when the release is
      // published, at the commit being released: the one a workflow checked
      // out, or HEAD here.
      target_commitish: commit || undefined,
      name: `Rutba Office ${tag.replace(/^v/, '')}`,
      body: notes,
      draft: true,
      // Beta: every release is a pre-release until the owner says otherwise.
      prerelease,
    });
    log(`\ncreated draft release #${release.id}`);
  }

  // 2. Every file, once, whole.
  const there = new Map((await client.listAssets(release)).map((a) => [a.name, a]));
  for (const name of assets) {
    const have = there.get(name);
    if (have && have.state === 'uploaded' && have.size === local.get(name).size) {
      log(`= ${name} (already uploaded)`);
      continue;
    }
    if (have) {
      log(`× ${name} (${have.state}, ${have.size} bytes) removed to upload again`);
      await client.deleteAsset(release, have);
    }
    log(`↑ ${name} (${(local.get(name).size / 1048576).toFixed(1)} MB)`);
    await client.uploadAsset(release, name, path.join(dir, name), local.get(name).size);
  }

  // 3. The check, against what GitHub now holds.
  const problems = [];
  const listed = await client.listAssets(release);
  const uploaded = new Map(listed.filter((a) => a.state === 'uploaded').map((a) => [a.name, a.size]));
  for (const name of assets) {
    if (!uploaded.has(name)) problems.push(`${name} is not in the release`);
    else if (uploaded.get(name) !== local.get(name).size) problems.push(`${name} is ${uploaded.get(name)} bytes in the release and ${local.get(name).size} here`);
  }
  for (const name of assets.filter((a) => /^latest.*\.yml$/.test(a))) {
    const asset = listed.find((a) => a.name === name);
    if (!asset) continue;
    // The feed as an installed copy will read it: from the release, not from this disk.
    const text = (await client.downloadAsset(release, asset)).toString('utf8');
    problems.push(...checkFeed({ name, feed: parseFeed(text), tag, uploaded, local }));
  }
  if (problems.length) {
    for (const p of problems) log(`! ${p}`);
    throw new Error(`the release's files do not agree, so ${release.draft ? 'the draft was not published' : 'the release needs mending'}: ${problems.length} problem${problems.length === 1 ? '' : 's'}`);
  }
  log(`\nchecked: ${assets.length} files uploaded at their sizes; the update feed matches them`);

  // 4. Published: from here the updater's feed and the releases page show it.
  if (release.draft) {
    release = await client.updateRelease(release, { draft: false, tag_name: tag, ...(commit ? { target_commitish: commit } : {}) });
    log(`published release #${release.id}`);
  }
  log(`\n${release.html_url}`);
  return release;
}

/** GitHub's REST interface, with the token held here. */
export function githubClient(token, { owner = OWNER, repo = REPO } = {}) {
  async function api(url, { method = 'GET', body, headers = {}, raw, binary = false } = {}) {
    const response = await fetch(url, {
      method,
      headers: {
        authorization: `Bearer ${token}`,
        accept: binary ? 'application/octet-stream' : 'application/vnd.github+json',
        'x-github-api-version': '2022-11-28',
        'user-agent': 'rutba-office-release',
        ...headers,
      },
      body: raw ?? (body ? JSON.stringify(body) : undefined),
      duplex: raw ? 'half' : undefined,
      redirect: 'follow',
    });
    if (binary && response.ok) return Buffer.from(await response.arrayBuffer());
    const text = await response.text();
    let parsed = null;
    try {
      parsed = text ? JSON.parse(text) : null;
    } catch {
      parsed = { raw: text.slice(0, 400) };
    }
    if (!response.ok) {
      const message = parsed?.message || response.statusText;
      const detail = parsed?.errors ? `: ${JSON.stringify(parsed.errors)}` : '';
      throw new Error(`${method} ${url.replace(/\?.*/, '')} → ${response.status} ${message}${detail}`);
    }
    return parsed;
  }
  const base = `https://api.github.com/repos/${owner}/${repo}`;
  return {
    // Drafts are listed only to a token that can push, which is the one this runs with.
    listReleases: async () => {
      const all = [];
      for (let page = 1; page <= 10; page++) {
        const batch = await api(`${base}/releases?per_page=100&page=${page}`);
        all.push(...batch);
        if (batch.length < 100) break;
      }
      return all;
    },
    createRelease: (body) => api(`${base}/releases`, { method: 'POST', body }),
    updateRelease: (release, body) => api(`${base}/releases/${release.id}`, { method: 'PATCH', body }),
    listAssets: (release) => api(`${base}/releases/${release.id}/assets?per_page=100`),
    deleteAsset: (release, asset) => api(`${base}/releases/assets/${asset.id}`, { method: 'DELETE' }),
    downloadAsset: (release, asset) => api(`${base}/releases/assets/${asset.id}`, { binary: true }),
    uploadAsset: (release, name, file, size) =>
      api(`${release.upload_url.replace(/\{.*$/, '')}?name=${encodeURIComponent(name)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/octet-stream', 'content-length': String(size) },
        raw: fs.createReadStream(file),
      }),
  };
}

/**
 * The token git already holds for github.com, or, where git holds none
 * because the credential lives in the environment (a build box whose
 * proxy signs the requests, a runner with GH_TOKEN set), that.
 */
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

/**
 * The release notes.
 *
 * From `docs/releases/<tag>.md` when there is one, because notes belong with
 * the release they describe rather than inside the script that uploads it:
 * a hard-coded block is one that still says 1.0.0 when 1.4 goes out.
 */
function notesFor(forTag) {
  const file = path.join(root, 'docs', 'releases', `${forTag}.md`);
  if (fs.existsSync(file)) return fs.readFileSync(file, 'utf8');
  throw new Error(
    `No notes for ${forTag}. Write docs/releases/${forTag}.md: a release with no notes is a file nobody knows whether to install.`
  );
}

async function main() {
  const tag = process.argv[2] || `v${JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version}`;
  // The file is wrapped for editors; the release page breaks a line wherever
  // the body does, so each paragraph goes up as one line (release-notes.js).
  const notes = unwrapNotes(notesFor(tag));
  await publishRelease({
    client: githubClient(credential()),
    dir: path.join(root, 'apps', 'desktop', 'release'),
    tag,
    notes,
    commit: process.env.GITHUB_SHA || null,
    // Beta: every release is a pre-release until the owner says otherwise
    // (RUTBA_RELEASE_FINAL=1 publishes a full release).
    prerelease: process.env.RUTBA_RELEASE_FINAL !== '1',
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    console.error(`\nfailed: ${err.message}`);
    process.exit(1);
  });
}
