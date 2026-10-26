// The release builds, after the renderer's bundle (npm run dist):
//
//   node build/dist.js [--win|--mac|--linux]
//
// The packs are made first (build/make-app-packs.js, signed with the key it is
// given), since an app that carries its code in its pack leaves that code out
// of the core (build/packs/moved.json, read by electron-builder.config.cjs).
//
// On Windows electron-builder then runs twice: the installer, which writes the
// packs a person ticks, and the portable build, which has no installer to
// choose in and so carries every pack (RUTBA_ALL_PACKS=1). The macOS disk
// image and the Linux packages carry every pack the same way.

import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

/** One step, from the app's folder; a failed step ends the build. */
function step(args, env = {}) {
  console.log(`\n> ${args.map((a) => path.basename(a)).join(' ')}`);
  const r = spawnSync(process.execPath, args, { cwd: APP, stdio: 'inherit', env: { ...process.env, ...env } });
  if (r.status !== 0) process.exit(r.status ?? 1);
}

const flag = process.argv.slice(2).find((a) => ['--win', '--mac', '--linux'].includes(a));
const platform = flag ? flag.slice(2) : { win32: 'win', darwin: 'mac' }[process.platform] || 'linux';
const builder = require.resolve('electron-builder/cli.js');
const config = ['--config', 'electron-builder.config.cjs', '--publish', 'never'];

step([path.join(APP, 'build', 'make-app-packs.js'), '--release', 'release']);
if (platform === 'win') {
  step([builder, ...config, '--win', 'nsis'], { RUTBA_ALL_PACKS: '' });
  step([builder, ...config, '--win', 'portable'], { RUTBA_ALL_PACKS: '1' });
} else {
  step([builder, ...config, `--${platform}`], { RUTBA_ALL_PACKS: '1' });
}
