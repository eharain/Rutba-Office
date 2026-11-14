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
//   4. left as a draft, which is what a build does here. Then, on the release
//      PC and only there, `node tools/publish-release.js <tag> --sign` with
//      RUTBA_UPDATE_KEY_FILE set reads each installer back from the draft,
//      holds it to the feed, signs it, and publishes (signAndPublish). The key
//      is not in Actions, so nothing can be signed there. Signing reads what
//      the draft holds, so a token able to change workflows could still put
//      files in a draft: what is published is checked against the commit, the
//      uploader and the feed, not against who ran the workflow. Nothing is
//      signed or published while a language the release leaves out of the
//      core (the contract's languages.leftCore) has no pack at this version
//      in the published packages feed, its signature checked
//      (packagesFeedProblems).
//
// A release already published is never changed: what is missing is uploaded to
// it and it is checked the same way, and a file there that is not the one built
// here stops the run, since people are downloading it. A new build is a new version.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { unwrapNotes, mentionsIn } from './release-notes.js';
import { UPDATE_KEYS, signedWords, signatureHolds } from '../apps/desktop/main/update-signature.js';
import { contractLeftCore, CONTRACT } from '../apps/desktop/build/write-left-core.js';
import { verifyFeed, acceptFeed } from '../packages/office-shell/src/electron/addons-compat.js';
import { releaseUrl } from '../apps/desktop/main/addons-download.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OWNER = 'eharain';
const REPO = 'Rutba-Office';

/** The files of this version in a build directory, in the order they go up. */
export function releaseAssets(names, tag) {
  const version = tag.replace(/^v/, '');
  return names
    // A .pack is an app Add or remove apps fetches from this release.
    .filter((name) => /\.(exe|dmg|zip|AppImage|deb|blockmap|pack)$/.test(name) || /^latest.*\.yml$/.test(name) || /-sbom\.cdx\.json$/.test(name))
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
 * shape electron-builder writes, which is all this has to read, and the one
 * flow list build/write-left-core.js adds (leftCore: ['ur', 'ar']).
 */
export function parseFeed(text) {
  const feed = { files: [] };
  let inFiles = false;
  let current = null;
  const scalar = (v) => {
    const s = v.trim();
    if (/^\[.*\]$/.test(s)) return s.slice(1, -1).split(',').map((x) => x.trim()).filter(Boolean).map(scalar);
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
 * `leftCore`, when given, is the contract's: the feed must say the same
 * languages left the core (build/write-left-core.js, D-2026-11-13-36).
 */
export function checkFeed({ name, feed, tag, uploaded, local, leftCore = null }) {
  const problems = [];
  if (Array.isArray(leftCore)) {
    const said = Array.isArray(feed.leftCore) ? [...new Set(feed.leftCore.map(String))].sort() : null;
    const want = [...new Set(leftCore)].sort();
    if (!said) problems.push(`${name} does not say which languages left the core (leftCore), so an older Office could not hold this update for a person in one of them`);
    else if (said.join(',') !== want.join(',')) problems.push(`${name} says the languages that left the core are [${said.join(', ')}], and the contract says [${want.join(', ')}]`);
  }
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

/** The packages keys of the add-ons contract: the keys a packages feed is signed with. */
const CONTRACT_JSON = JSON.parse(fs.readFileSync(CONTRACT, 'utf8'));
export const PACKAGES_KEYS = Object.freeze([...(CONTRACT_JSON.keys?.packages || [])]);

/** The packages feed for a version, as an installed copy fetches it: feed.json and feed.json.sig of that release. */
export async function fetchPackagesFeed(version, { fetchImpl = fetch } = {}) {
  const text = async (url) => {
    const r = await fetchImpl(url, { headers: { 'user-agent': 'Rutba-Office' }, redirect: 'follow', signal: AbortSignal.timeout(60_000) });
    if (!r.ok) throw new Error(`${url} answered ${r.status}`);
    return r.text();
  };
  const base = releaseUrl(version);
  const [feed, sig] = await Promise.all([text(`${base}feed.json`), text(`${base}feed.json.sig`)]);
  return { feed: JSON.parse(feed), sig: JSON.parse(sig) };
}

/**
 * What stops a release that leaves languages out of the core, as sentences;
 * empty when nothing does. The published packages feed for its version must
 * be signed with a packages key, be one Office itself takes (acceptFeed:
 * the contract's format, that version, a state and its lists), and list a pack
 * (lang-<tag>) at that version, not withdrawn or retired, for every language
 * in `leftCore`: else a person in one of them who updates is held, or with
 * "Update anyway" left with no words (D-2026-11-13-36). `packagesFeed` gives
 * { feed, sig } for a version (fetchPackagesFeed, or a test's stand-in).
 */
export async function packagesFeedProblems({ version, leftCore, packagesFeed = fetchPackagesFeed, keys = PACKAGES_KEYS }) {
  const tags = [...new Set(leftCore || [])].sort();
  if (!tags.length) return [];
  let got;
  try {
    got = await packagesFeed(version);
  } catch (err) {
    return [`the packages feed for v${version} could not be had (${err?.message || err}), so the packs for ${tags.join(', ')} cannot be shown to be published`];
  }
  const { feed, sig } = got || {};
  if (!feed || !verifyFeed(feed, sig, keys)) return [`the packages feed for v${version} is not signed with a packages key of the contract`];
  // Taken as an installed copy takes it (acceptFeed): its format, its version, its state and lists, its sequence.
  const taken = acceptFeed(feed, { office: version, contract: CONTRACT_JSON });
  if (!taken.ok) return [`the packages feed for v${version} is one Rutba Office ${version} would refuse: ${taken.why}`];
  const packs = Array.isArray(feed.packs) ? feed.packs : [];
  const retired = new Set((Array.isArray(feed.retired) ? feed.retired : []).map((r) => r?.id));
  const problems = [];
  for (const tag of tags) {
    const id = `lang-${tag}`;
    const p = packs.find((q) => q?.id === id);
    if (!p) problems.push(`the packages feed for v${version} lists no ${id}, and ${tag} leaves the core in this release`);
    else if (retired.has(id)) problems.push(`the packages feed for v${version} retires ${id}, and ${tag} leaves the core in this release`);
    else if (p.withdrawn) problems.push(`${id} is withdrawn in the packages feed for v${version}`);
    else if (p.version !== version) problems.push(`${id} in the packages feed for v${version} is at ${p.version}`);
  }
  return problems;
}

/**
 * Whether a file already in the release is this one. The size says little: a
 * rebuilt feed is the same length with another hash in it, and left there it
 * fails every copy that updates. A feed is small enough to read; GitHub gives
 * the digest of the rest, where it has one.
 */
async function isTheFileBuilt(client, release, have, built) {
  if (have.state !== 'uploaded' || have.size !== built.size) return false;
  if (/^latest.*\.yml$/.test(have.name)) return sha512Of(await client.downloadAsset(release, have)) === built.sha512;
  const digest = /^sha256:([0-9a-f]{64})$/i.exec(have.digest || '');
  // No digest to hold it to: the same size proves nothing, so it goes up again.
  return digest ? digest[1].toLowerCase() === built.sha256 : false;
}

/**
 * The release, made and published as the header says. `client` is GitHub
 * (githubClient below), or the engine tests' stand-in for it.
 */
export async function publishRelease({ client, dir, tag, notes, commit = null, prerelease = false, leaveDraft = false, leftCore = null, packagesFeed = fetchPackagesFeed, packagesKeys = PACKAGES_KEYS, log = console.log }) {
  if (!fs.existsSync(dir)) throw new Error(`nothing built: ${dir} does not exist`);
  const assets = releaseAssets(fs.readdirSync(dir), tag);
  if (!assets.length) throw new Error('no release artefacts found: run npm run dist first');
  if (!assets.some((a) => /^latest.*\.yml$/.test(a))) throw new Error('there is no update feed (latest.yml) beside the installers, so an installed copy would never find this release');

  log(`tag       ${tag}`);
  log(`artefacts ${assets.length}`);
  const local = new Map();
  for (const name of assets) {
    const bytes = fs.readFileSync(path.join(dir, name));
    local.set(name, { size: bytes.length, sha512: sha512Of(bytes), sha256: crypto.createHash('sha256').update(bytes).digest('hex') });
    log(`  ${name.padEnd(44)} ${(bytes.length / 1048576).toFixed(2)} MB`);
  }

  // 1. The draft: one an earlier run left, or a new one.
  let release = (await client.listReleases()).find((r) => r.tag_name === tag) || null;
  if (release) {
    log(`\nrelease already exists (#${release.id}, ${release.draft ? 'a draft' : 'published'}), reusing it`);
    // A draft left by another build tags this build's commit, not that one's.
    if (release.draft && release.body !== notes) release = await client.updateRelease(release, { body: notes });
    if (release.draft && commit && release.target_commitish !== commit) {
      // Every file of the other build goes first: a run that stops after this must not leave its files under this commit.
      for (const a of await client.listAssets(release)) await client.deleteAsset(release, a);
      release = await client.updateRelease(release, { target_commitish: commit });
    }
  } else {
    release = await client.createRelease({
      tag_name: tag,
      // The tag, if it is not pushed yet, is made when the release is
      // published, at the commit being released: the one a workflow checked
      // out, or HEAD here.
      target_commitish: commit || undefined,
      name: `Rutba Office ${tag.replace(/^v/, '')}`,
      body: notes,
      draft: true,
      prerelease,
    });
    log(`\ncreated draft release #${release.id}`);
  }

  // 2. Every file, once, whole.
  const there = new Map((await client.listAssets(release)).map((a) => [a.name, a]));
  // A draft holds this build and its signatures, and nothing an earlier build made and this one did not.
  if (release.draft) {
    for (const [name, asset] of there) {
      if (assets.includes(name) || (name.endsWith('.sig') && assets.includes(name.slice(0, -4)))) continue;
      log(`× ${name} (not made by this build) removed`);
      await client.deleteAsset(release, asset);
      there.delete(name);
    }
  }
  // Held to the build before anything moves: a published release that disagrees is left exactly as it is.
  const same = new Set();
  for (const name of assets) {
    const have = there.get(name);
    if (!have) continue;
    if (await isTheFileBuilt(client, release, have, local.get(name))) same.add(name);
    else if (!release.draft) throw new Error(`${name} in the published release ${tag} is not the file built here, and a published release is never changed. Make a new version instead.`);
  }
  for (const name of assets) {
    const have = there.get(name);
    if (same.has(name)) {
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
    problems.push(...checkFeed({ name, feed: parseFeed(text), tag, uploaded, local, leftCore }));
  }
  // Published here only when not left as a draft: then the packs that take the languages out of the core must be out.
  if (release.draft && !leaveDraft) problems.push(...await packagesFeedProblems({ version: tag.replace(/^v/, ''), leftCore, packagesFeed, keys: packagesKeys }));
  if (problems.length) {
    for (const p of problems) log(`! ${p}`);
    throw new Error(`the release's files do not agree, so ${release.draft ? 'the draft was not published' : 'the release needs mending'}: ${problems.length} problem${problems.length === 1 ? '' : 's'}`);
  }
  log(`\nchecked: ${assets.length} files uploaded at their sizes; the update feed matches them`);

  // 4. Published: from here the updater's feed and the releases page show it.
  if (release.draft && leaveDraft) {
    log(`left as draft #${release.id}: it is signed and published with --sign on the release PC`);
  } else if (release.draft) {
    release = await client.updateRelease(release, { draft: false, tag_name: tag, ...(commit ? { target_commitish: commit } : {}) });
    log(`published release #${release.id}`);
  }
  log(`\n${release.html_url}`);
  return release;
}

/** Who uploads what a release carries: the release workflow, never a person's token. */
const BUILT_BY = 'github-actions[bot]';
const isBuilt = (name) => /\.exe$|\.pack$|^latest.*\.yml$/i.test(name);

/**
 * Signs what is in the draft and publishes it. `sign` turns the words into
 * the text of a .sig file; the key stays with whoever calls this. Nothing is
 * taken from this disk: what is signed is what people will download.
 */
export async function signAndPublish({ client, tag, sign, commit = null, keys = UPDATE_KEYS, builtBy = BUILT_BY, leftCore = null, packagesFeed = fetchPackagesFeed, packagesKeys = PACKAGES_KEYS, log = console.log }) {
  const version = tag.replace(/^v/, '');
  const named = (await client.listReleases()).filter((r) => r.tag_name === tag);
  if (named.length > 1) throw new Error(`there are ${named.length} releases for ${tag}: one of them has to go before any is signed`);
  let release = named[0] || null;
  if (!release) throw new Error(`there is no release for ${tag} to sign`);
  // What is signed is what the gate passed: the commit, and files the workflow itself built from it.
  if (release.draft && commit && release.target_commitish !== commit) throw new Error(`the draft of ${tag} was built from ${release.target_commitish} and this copy is at ${commit}: nothing was signed`);
  const listed = await client.listAssets(release);
  const notBuilt = listed.filter((a) => isBuilt(a.name) && a.uploader?.login !== builtBy).map((a) => a.name);
  if (notBuilt.length) throw new Error(`${notBuilt.join(', ')} in the draft of ${tag} did not come from the release workflow: nothing was signed`);
  const byName = new Map(listed.map((a) => [a.name, a]));
  const uploaded = new Map(listed.filter((a) => a.state === 'uploaded').map((a) => [a.name, a.size]));
  const feedAsset = byName.get('latest.yml');
  if (!feedAsset) throw new Error('the release has no latest.yml');
  const feed = parseFeed((await client.downloadAsset(release, feedAsset)).toString('utf8'));
  const problems = checkFeed({ name: 'latest.yml', feed, tag, uploaded, local: new Map(), leftCore });
  if (problems.length) throw new Error(`the release's files do not agree: ${problems.join('; ')}`);
  // The languages the release leaves out of the core have their packs out, signed, before anything is signed.
  const lacking = await packagesFeedProblems({ version, leftCore, packagesFeed, keys: packagesKeys });
  if (lacking.length) throw new Error(`the packages are not out for this release, so nothing was signed: ${lacking.join('; ')}`);
  const installers = feed.files.filter((f) => /\.exe$/i.test(String(f.url)));
  if (!installers.length) throw new Error('latest.yml names no installer');
  // An app pack is signed with the same key: the key the release workflow holds is not enough to put code on a person's computer.
  const packs = listed.filter((a) => /\.pack$/i.test(a.name));
  // A good signature stays: taking it down to put up another leaves a moment when a published file has none.
  const stillGood = async (name, words) => {
    if (!byName.has(name)) return false;
    if (!signatureHolds(words, (await client.downloadAsset(release, byName.get(name))).toString('utf8'), keys)) return false;
    log(`${name} is already good`);
    return true;
  };
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sign-'));
  try {
    for (const f of installers) {
      const file = decodeURIComponent(String(f.url)).split('/').pop();
      const sha512 = sha512Of(await client.downloadAsset(release, byName.get(file)));
      if (sha512 !== f.sha512) throw new Error(`${file} in the release is not what latest.yml says it is`);
      const words = signedWords({ version, file, sha512 });
      if (await stillGood(`${file}.sig`, words)) continue;
      const text = sign(words);
      if (!signatureHolds(words, text, keys)) throw new Error(`this key is not one the program trusts, so installed copies would refuse ${file}`);
      const name = `${file}.sig`;
      if (byName.has(name)) await client.deleteAsset(release, byName.get(name));
      const out = path.join(scratch, name);
      fs.writeFileSync(out, `${text}\n`);
      await client.uploadAsset(release, name, out, fs.statSync(out).size);
      log(`signed ${file}`);
    }
    for (const asset of packs) {
      const words = signedWords({ version, file: asset.name, sha512: sha512Of(await client.downloadAsset(release, asset)) });
      if (await stillGood(`${asset.name}.sig`, words)) continue;
      const text = sign(words);
      if (!signatureHolds(words, text, keys)) throw new Error('this key is not one the program trusts, so installed copies would refuse the app packs');
      const name = `${asset.name}.sig`;
      if (byName.has(name)) await client.deleteAsset(release, byName.get(name));
      const out = path.join(scratch, name);
      fs.writeFileSync(out, `${text}\n`);
      await client.uploadAsset(release, name, out, fs.statSync(out).size);
      log(`signed ${asset.name}`);
    }
  } finally {
    fs.rmSync(scratch, { recursive: true, force: true });
  }
  if (release.draft) {
    release = await client.updateRelease(release, { draft: false, tag_name: tag });
    log(`published release #${release.id}`);
  }
  log(release.html_url);
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
/** The commit the note's Gate section says was checked, or null. */
export function gatedCommitOf(note) {
  const body = String(note || '').replace(/\r\n/g, '\n');
  const at = body.search(/^## Gate\n/m);
  if (at < 0) return null;
  return /^Commit checked: ([0-9a-f]{40})\s*$/m.exec(body.slice(at))?.[1] || null;
}

/**
 * Why this release may not go up, or null. What is published must be what the
 * gate checked: the only files that may differ from the commit it checked are
 * the ones its record writes, as tools/gate.js dirtyForRecord has them: the
 * README, the floors, the release note, and the numbers given to checks seen
 * for the first time (12 Nov: a run that gave one a number was refused here).
 */
export function gateProblem({ note, tag, changed }) {
  const commit = gatedCommitOf(note);
  if (!commit) return `docs/releases/${tag}.md has no Gate section naming the commit that was checked. Run the gate with --record.`;
  if (changed === null) return `Commit ${commit}, which the gate checked, is not in this checkout.`;
  const allowed = new Set(['README.md', 'tools/gate-floors.json', 'tools/check-ids.json', `docs/releases/${tag}.md`]);
  const extra = changed.filter((f) => !allowed.has(f));
  return extra.length ? `${extra.length} file(s) changed since the gate checked ${commit.slice(0, 8)}: ${extra.slice(0, 5).join(', ')}. Run the gate again.` : null;
}

function changedSince(commit) {
  try {
    return execFileSync('git', ['diff', '--name-only', commit, 'HEAD'], { cwd: root, encoding: 'utf8' }).split('\n').map((l) => l.trim()).filter(Boolean);
  } catch {
    return null;
  }
}

function notesFor(forTag) {
  const file = path.join(root, 'docs', 'releases', `${forTag}.md`);
  if (fs.existsSync(file)) return fs.readFileSync(file, 'utf8');
  throw new Error(
    `No notes for ${forTag}. Write docs/releases/${forTag}.md: a release with no notes is a file nobody knows whether to install.`
  );
}

async function main() {
  const args = process.argv.slice(2);
  const signing = args.includes('--sign');
  const packaged = `v${JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version}`;
  const tag = args.find((a) => !a.startsWith('--')) || packaged;
  // A build is of one version: a tag that says another would go up with the wrong files.
  if (tag !== packaged) throw new Error(`the tag is ${tag} and package.json says ${packaged}: they must be the same`);
  // The file is wrapped for editors; the release page breaks a line wherever
  // the body does, so each paragraph goes up as one line (release-notes.js).
  const written = notesFor(tag);
  const commit = gatedCommitOf(written);
  const refused = gateProblem({ note: written, tag, changed: commit ? changedSince(commit) : [] });
  if (refused) throw new Error(refused);
  const notes = unwrapNotes(written);
  // A name after an @ is a GitHub account: the release page would list it
  // as a contributor and notify it (release-notes.js mentionsIn).
  const mentioned = mentionsIn(notes);
  if (mentioned.length) {
    throw new Error(`docs/releases/${tag}.md mentions ${mentioned.join(', ')} on GitHub, which would list them as the release's contributors and notify them. Put each in backticks, as code.`);
  }
  if (signing) {
    const keyFile = process.env.RUTBA_UPDATE_KEY_FILE;
    if (!keyFile || !fs.existsSync(keyFile)) throw new Error('Set RUTBA_UPDATE_KEY_FILE to the update key on the release PC.');
    const key = crypto.createPrivateKey(fs.readFileSync(keyFile));
    const head = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
    let tagged = null;
    try { tagged = execFileSync('git', ['rev-parse', '--verify', '--quiet', `${tag}^{commit}`], { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim(); } catch { /* the release makes the tag when it is published */ }
    if (tagged && tagged !== head) throw new Error(`${tag} is at ${tagged} and this copy is at ${head}: nothing was signed`);
    // The tag on GitHub is the one the release attaches to, and it may not be here. Its peeled line comes last.
    const remote = execFileSync('git', ['ls-remote', 'origin', `refs/tags/${tag}`, `refs/tags/${tag}^{}`], { cwd: root, encoding: 'utf8' }).trim().split('\n').filter(Boolean).pop()?.split(/\s/)[0];
    if (remote && remote !== head) throw new Error(`${tag} on GitHub is at ${remote} and this copy is at ${head}: nothing was signed`);
    await signAndPublish({ client: githubClient(credential()), tag, commit: head, leftCore: contractLeftCore(), sign: (words) => crypto.sign(null, Buffer.from(words, 'utf8'), key).toString('base64') });
    return;
  }
  await publishRelease({
    leaveDraft: true,
    client: githubClient(credential()),
    dir: path.join(root, 'apps', 'desktop', 'release'),
    tag,
    notes,
    commit: process.env.GITHUB_SHA || null,
    leftCore: contractLeftCore(),
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    console.error(`\nfailed: ${err.message}`);
    process.exit(1);
  });
}
