// How many people have downloaded Rutba Office.
//
// GitHub counts this already and has since the release was published — there is
// nothing to add to the application and nothing to ask anyone's permission for,
// because it is a count of requests GitHub served, not of people.
//
//   node tools/downloads.js            every release, newest first
//   node tools/downloads.js --total    just the number
//   node tools/downloads.js --json     for whatever wants to draw it
//
// Authentication is optional: the releases endpoint is public. With a token the
// rate limit goes from 60 requests an hour to 5,000, which only matters if this
// is run on a schedule. The token is read the way the publisher reads it — from
// the credential helper git already uses — so there is nothing to paste.

import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const run = promisify(execFile);
const REPO = 'eharain/Rutba-Office';

/** The token git already holds for github.com, or nothing. */
async function token() {
  try {
    const { stdout } = await run('git', ['credential', 'fill'], {
      input: 'protocol=https\nhost=github.com\n\n',
      windowsHide: true,
      // Without these the helper opens a dialog and waits, which turns a
      // read-only report into a hung terminal on any machine that has no
      // credential stored for github.com.
      env: { ...process.env, GIT_TERMINAL_PROMPT: '0', GCM_INTERACTIVE: 'never' },
      timeout: 5000,
    });
    return /password=(.+)/.exec(stdout)?.[1]?.trim() || null;
  } catch {
    return null;
  }
}

async function releases() {
  const auth = await token();
  const response = await fetch(`https://api.github.com/repos/${REPO}/releases?per_page=100`, {
    headers: {
      Accept: 'application/vnd.github+json',
      'User-Agent': 'rutba-office-downloads',
      ...(auth ? { Authorization: `Bearer ${auth}` } : {}),
    },
  });
  if (!response.ok) throw new Error(`GitHub answered ${response.status} ${response.statusText}`);
  return response.json();
}

/** Which platform an asset is for, from the name the builder gave it. */
function platformOf(name) {
  if (/\.(exe|msi)$/i.test(name)) return 'Windows';
  if (/\.(dmg|pkg)$/i.test(name)) return 'macOS';
  if (/\.(AppImage|deb|rpm|snap|tar\.gz)$/i.test(name)) return 'Linux';
  if (/\.(yml|yaml|blockmap)$/i.test(name)) return 'update metadata';
  return 'other';
}

const list = await releases();
const wantJson = process.argv.includes('--json');
const wantTotal = process.argv.includes('--total');

const report = list.map((release) => {
  const assets = (release.assets || [])
    .map((a) => ({ name: a.name, platform: platformOf(a.name), downloads: a.download_count, size: a.size }))
    // The update feed is fetched by every installed copy on a timer, so counting
    // it as a download would flatter the number by an order of magnitude.
    .filter((a) => a.platform !== 'update metadata');
  return {
    tag: release.tag_name,
    published: release.published_at,
    prerelease: release.prerelease,
    downloads: assets.reduce((n, a) => n + a.downloads, 0),
    assets,
  };
});

const total = report.reduce((n, r) => n + r.downloads, 0);

if (wantJson) {
  console.log(JSON.stringify({ repo: REPO, total, releases: report }, null, 2));
} else if (wantTotal) {
  console.log(String(total));
} else {
  if (!report.length) {
    console.log('No releases published yet.');
  }
  for (const release of report) {
    const when = release.published ? new Date(release.published).toLocaleDateString() : 'unpublished';
    console.log(`\n${release.tag}${release.prerelease ? ' (pre-release)' : ''} — ${when} — ${release.downloads.toLocaleString()} downloads`);
    const byPlatform = new Map();
    for (const a of release.assets) byPlatform.set(a.platform, (byPlatform.get(a.platform) || 0) + a.downloads);
    for (const [platform, n] of [...byPlatform].sort((a, b) => b[1] - a[1])) {
      console.log(`  ${platform.padEnd(10)} ${String(n).padStart(7)}`);
    }
    for (const a of release.assets.filter((x) => x.downloads > 0)) {
      console.log(`    ${a.name} — ${a.downloads.toLocaleString()}`);
    }
  }
  console.log(`\n${total.toLocaleString()} downloads in total across ${report.length} release${report.length === 1 ? '' : 's'}.`);
  console.log('This counts requests GitHub served. It does not identify anyone, and it does not count copies still running.');
}
