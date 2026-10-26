// The command AI Tools shows an AI app (main/mcp/command.js), the start file
// an installed copy names (resources/mcp/rutba-office-mcp.cjs), and the
// profile the tools then work in (main/mcp/profile-dir.js). Until 2.18 the
// lines named a file inside app.asar, which an AI app checking the path first
// could not find (owner, 2026-10-26: "when i passed this mcp cli the tool
// complained its not there").

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { mcpCommand, INSTALLED } from '../apps/desktop/main/mcp/command.js';
import { profileDir, baseDir } from '../apps/desktop/main/mcp/profile-dir.js';
import { makeProfile } from '@rutba/office-shell/profiles';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const desktop = path.join(root, 'apps', 'desktop');
const launcherFile = path.join(desktop, 'resources', 'mcp', 'rutba-office-mcp.cjs');
const require = createRequire(import.meta.url);
const temp = (name) => fs.mkdtempSync(path.join(os.tmpdir(), name));

test('an installed copy names the start file beside the archive, with the profile, run as Node', () => {
  const resources = path.join('C:', 'Programs', 'Rutba Office', 'resources');
  const c = mcpCommand({ execPath: 'C:\\Programs\\Rutba Office\\Rutba Office.exe', packaged: true, resourcesPath: resources, profile: { id: 'a1b2c3d4', folder: 'x' } });
  assert.equal(c.command, 'C:\\Programs\\Rutba Office\\Rutba Office.exe');
  assert.deepEqual(c.args, [path.join(resources, 'mcp', 'rutba-office-mcp.cjs'), '--profile', 'a1b2c3d4']);
  assert.deepEqual(c.env, { ELECTRON_RUN_AS_NODE: '1' });
  assert.ok(!c.args.some((a) => a.includes('app.asar')), 'no path inside the archive');
});

test('a copy run from the source names the tools in its own folder, and Default when nothing is settled', () => {
  const c = mcpCommand({ execPath: 'electron.exe', appDir: desktop, packaged: false });
  assert.deepEqual(c.args, [path.join(desktop, 'main', 'mcp', 'node.js'), '--profile', 'default']);
  assert.ok(fs.existsSync(c.args[0]));
});

test('a base folder named on the command line goes with the command, so the tools look where the run does', () => {
  const c = mcpCommand({ execPath: 'electron.exe', appDir: desktop, profile: { id: 'default', folder: 'D:\\checks\\base', baseNamed: true } });
  assert.equal(c.env.RUTBA_OFFICE_PROFILE, 'D:\\checks\\base');
  assert.equal(mcpCommand({ execPath: 'e', appDir: desktop, profile: { id: 'default', folder: 'f' } }).env.RUTBA_OFFICE_PROFILE, undefined);
});

test('a screens run shows the lines an installed copy shows, as docs/AI-TOOLS.md prints them', () => {
  const c = mcpCommand({ execPath: 'electron.exe', appDir: desktop, screens: true, profile: { id: 'a1b2c3d4', folder: 'f', baseNamed: true } });
  assert.equal(c.command, `${INSTALLED}\\Rutba Office.exe`);
  assert.deepEqual(c.args, [`${INSTALLED}\\resources\\mcp\\rutba-office-mcp.cjs`, '--profile', 'default']);
  assert.deepEqual(c.env, { ELECTRON_RUN_AS_NODE: '1' });
  const doc = fs.readFileSync(path.join(root, 'docs', 'AI-TOOLS.md'), 'utf8');
  assert.ok(doc.includes(JSON.stringify(c.args).replace(/,/g, ', ')), 'the guide prints the same args');
});

test('the start file finds the archive beside its own folder', () => {
  const { archiveEntry } = require(launcherFile);
  assert.equal(archiveEntry(path.join('C:', 'Rutba Office', 'resources', 'mcp')), path.join('C:', 'Rutba Office', 'resources', 'app.asar', 'main', 'mcp', 'node.js'));
  assert.equal(archiveEntry(), path.join(desktop, 'resources', 'app.asar', 'main', 'mcp', 'node.js'), 'by default from its own __dirname');
});

test('run from an install, the start file loads the tools beside it and hands them its words', () => {
  const install = temp('mcp-install-');
  try {
    const mcp = path.join(install, 'resources', 'mcp');
    const tools = path.join(install, 'resources', 'app.asar', 'main', 'mcp');
    fs.mkdirSync(mcp, { recursive: true });
    fs.mkdirSync(tools, { recursive: true });
    fs.copyFileSync(launcherFile, path.join(mcp, 'rutba-office-mcp.cjs'));
    fs.writeFileSync(path.join(install, 'resources', 'app.asar', 'package.json'), '{"type":"module"}');
    fs.writeFileSync(path.join(tools, 'node.js'), "process.stdout.write(JSON.stringify({ argv: process.argv.slice(2) }));\n");
    const out = JSON.parse(execFileSync(process.execPath, [path.join(mcp, 'rutba-office-mcp.cjs'), '--profile', 'work'], { encoding: 'utf8', timeout: 20000 }));
    assert.deepEqual(out.argv, ['--profile', 'work']);
  } finally {
    fs.rmSync(install, { recursive: true, force: true });
  }
});

test('the installer ships the start file beside the archive', async () => {
  const config = await require(path.join(desktop, 'electron-builder.config.cjs'))();
  const entry = (config.extraResources || []).find((e) => e.to === 'mcp');
  assert.ok(entry, 'an extraResources entry for mcp');
  assert.ok(fs.existsSync(path.join(desktop, entry.from, 'rutba-office-mcp.cjs')));
});

test('the tools work in the profile the command names, by id or name, and refuse one that is not there', () => {
  const appData = temp('mcp-appdata-');
  try {
    const env = { APPDATA: appData, XDG_CONFIG_HOME: appData };
    const base = baseDir(env);
    fs.mkdirSync(base, { recursive: true });
    assert.equal(profileDir({ argv: [], env }), base, 'no profile named: the base folder, as before profiles');
    assert.equal(profileDir({ argv: ['--profile', 'default'], env }), base, 'Default without a profiles list');
    const { profile } = makeProfile(base, { name: 'Work' });
    assert.equal(profileDir({ argv: ['--profile', profile.id], env }), path.join(base, 'Profiles', profile.id));
    assert.equal(profileDir({ argv: ['--profile=work'], env }), path.join(base, 'Profiles', profile.id), 'by name, any case');
    assert.throws(() => profileDir({ argv: ['--profile', 'gone'], env }), /no profile "gone"/);
    assert.equal(profileDir({ argv: ['--profile', 'gone'], env: { ...env, RUTBA_OFFICE_PROFILE: 'D:\\x' } }), 'D:\\x', 'the folder named outright wins');
  } finally {
    fs.rmSync(appData, { recursive: true, force: true });
  }
});

// The real thing: Electron run as Node reads the tools out of an asar archive
// through the start file beside it, as an installed copy does.
test('Electron run as Node starts the tools from inside an archive through the start file', async (t) => {
  let asar;
  try {
    asar = require('@electron/asar');
  } catch {
    t.skip('@electron/asar is not installed');
    return;
  }
  const electron = createRequire(path.join(desktop, 'package.json'))('electron');
  const work = temp('mcp-asar-');
  try {
    const src = path.join(work, 'src');
    fs.mkdirSync(path.join(src, 'main', 'mcp'), { recursive: true });
    fs.writeFileSync(path.join(src, 'package.json'), '{"name":"probe","type":"module"}');
    fs.writeFileSync(path.join(src, 'main', 'mcp', 'said.js'), "export const said = 'from the archive';\n");
    fs.writeFileSync(path.join(src, 'main', 'mcp', 'node.js'), "import { said } from './said.js';\nprocess.stdout.write(JSON.stringify({ said, argv: process.argv.slice(2) }));\n");
    const resources = path.join(work, 'resources');
    fs.mkdirSync(path.join(resources, 'mcp'), { recursive: true });
    await asar.createPackage(src, path.join(resources, 'app.asar'));
    fs.copyFileSync(launcherFile, path.join(resources, 'mcp', 'rutba-office-mcp.cjs'));
    const env = { ...process.env, ELECTRON_RUN_AS_NODE: '1' };
    const out = JSON.parse(execFileSync(electron, [path.join(resources, 'mcp', 'rutba-office-mcp.cjs'), '--profile', 'a1b2c3d4'], { env, encoding: 'utf8', timeout: 60000 }));
    assert.deepEqual(out, { said: 'from the archive', argv: ['--profile', 'a1b2c3d4'] });
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
});
