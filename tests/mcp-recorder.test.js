// The Recorder's AI tools (2.22, D-2026-10-30-06): Off until the person sets
// the Recorder to record in AI Tools; a recording goes only into a folder the
// tools may make files in; start and stop are asked of the Recorder window
// (office-shell recorder-remote.js), which answers with the file; the log
// names the Recorder and the step. Going live is never a tool.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createReachTools, REACH_TOOL_LIST } from '../apps/desktop/main/mcp/reach-tools.js';
import { FOLDERS_KEY, resolveIn, foldersOf } from '../apps/desktop/main/mcp/access.js';
import { RECORDER_KEY, stepsFor, stepOf } from '../apps/desktop/main/mcp/reach.js';
import { toolOffered, LOADED_KEY } from '../packages/office-shell/src/pack-tools.js';

const textOf = (r) => (r?.content || []).map((c) => c.text).join('\n');

function memorySettings(initial = {}) {
  const s = { ...initial };
  return { get: (k, f) => (k in s ? s[k] : f), set: (k, v) => { s[k] = v; } };
}

/** The Recorder window as the tools reach it: what it was asked, and its answers. */
function fakeRecorder() {
  const asked = [];
  let phase = 'setup';
  return {
    asked,
    ask: async (op, args = {}) => {
      asked.push({ op, args });
      if (op === 'sources') return [{ id: 'window:1:0', name: 'Rutba Office: Sales', kind: 'window' }, { id: 'screen:0:0', name: 'Entire screen', kind: 'screen' }];
      if (op === 'state') return { phase, seconds: phase === 'recording' ? 5 : 0, disk: { level: 'ok', minutes: 600 } };
      if (op === 'start') { phase = 'recording'; return { recording: true, path: path.join(args.folder, 'Recording 1.webm'), source: 'Rutba Office: Sales' }; }
      if (op === 'stop') { phase = 'setup'; return { path: 'Recording 1.webm', seconds: 5, bytes: 123456 }; }
      throw new Error(`no ${op}`);
    },
  };
}

function setup(step = 'off') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-rec-'));
  const made = path.join(dir, 'made');
  const read = path.join(dir, 'read');
  fs.mkdirSync(made);
  fs.mkdirSync(read);
  const settings = memorySettings({ [RECORDER_KEY]: step, [FOLDERS_KEY]: [made, { path: read, step: 'read' }] });
  const recorder = fakeRecorder();
  const tools = createReachTools({
    settings,
    recorder,
    confirm: { ask: async () => ({}) },
    files: {
      readable: (given) => resolveIn(foldersOf({ [FOLDERS_KEY]: settings.get(FOLDERS_KEY) }), given),
      writable: (given) => {
        const list = foldersOf({ [FOLDERS_KEY]: settings.get(FOLDERS_KEY) }, 'prepare');
        return !given && list.length ? list[0] : resolveIn(list, given);
      },
    },
  });
  return { dir, made, read, settings, recorder, tools };
}

test('the Recorder is Off, or records, and nothing else', () => {
  assert.deepEqual(stepsFor('recorder'), ['off', 'prepare']);
  assert.equal(stepOf('recorder', 'act'), 'off');
  assert.equal(stepOf('recorder', undefined), 'off');
  const names = REACH_TOOL_LIST.map((t) => t.name);
  for (const n of ['recorder_windows', 'recorder_start', 'recorder_stop', 'recorder_state']) assert.ok(names.includes(n), n);
  assert.ok(!names.some((n) => /live|stream|broadcast/i.test(n)), 'going live is never a tool');
});

test('Off, every Recorder tool is refused, naming where its setting is, and nothing is asked of the window', async () => {
  const { tools, recorder } = setup('off');
  for (const name of ['recorder_windows', 'recorder_start', 'recorder_stop', 'recorder_state']) {
    const { result, log } = await tools.call(name, name === 'recorder_start' ? { source: 'window:1:0' } : {});
    assert.equal(result.isError, true, name);
    assert.match(textOf(result), /the Recorder: it is set to Off.*AI Tools, under Recorder/, name);
    assert.deepEqual(log, { kind: 'recorder', name: 'Recorder', step: 'prepare' });
  }
  assert.equal(recorder.asked.length, 0);
});

test('set to record: the windows listed, a recording started in a given folder, stopped with its file, length and size', async () => {
  const { tools, recorder, made } = setup('prepare');
  const listed = await tools.call('recorder_windows');
  assert.match(textOf(listed.result), /window\twindow:1:0\tRutba Office: Sales/);
  const started = await tools.call('recorder_start', { source: 'window:1:0', title: 'Sales', microphone: false });
  assert.ok(!started.result.isError, textOf(started.result));
  assert.deepEqual(started.log, { kind: 'recorder', name: 'Recorder', step: 'prepare' });
  const asked = recorder.asked.find((a) => a.op === 'start').args;
  assert.equal(asked.folder, made, 'the first folder it may make files in, when none is named');
  assert.equal(asked.sound, true);
  assert.equal(asked.microphone, false);
  assert.equal(asked.camera, false);
  assert.match(textOf((await tools.call('recorder_state')).result), /phase\trecording\nseconds\t5/);
  const stopped = await tools.call('recorder_stop');
  assert.match(textOf(stopped.result), /Kept Recording 1\.webm: 5 seconds, 123456 bytes\./);
});

test('a recording never goes into a folder the tools only read, or one not given', async () => {
  const { tools, recorder, read, dir } = setup('prepare');
  for (const folder of [read, path.join(dir, 'elsewhere')]) {
    const { result } = await tools.call('recorder_start', { source: 'window:1:0', folder });
    assert.equal(result.isError, true, folder);
  }
  assert.ok(!recorder.asked.some((a) => a.op === 'start'));
});

test('a camera scene turns the camera on; the tools are offered only with the Recorder installed', async () => {
  const { tools, recorder } = setup('prepare');
  await tools.call('recorder_start', { source: 'window:1:0', scene: 'split' });
  assert.equal(recorder.asked.find((a) => a.op === 'start').args.camera, true);
  assert.equal(toolOffered('recorder_start', { [LOADED_KEY]: ['ai', 'video'] }), false);
  assert.equal(toolOffered('recorder_start', { [LOADED_KEY]: ['ai', 'recorder'] }), true);
});

test('a copy without a Recorder lists none, and says so when asked', async () => {
  const tools = createReachTools({ settings: memorySettings({ [RECORDER_KEY]: 'prepare' }), confirm: {}, files: {} });
  assert.ok(!tools.things().some((x) => x.kind === 'recorder'));
  const { result } = await tools.call('recorder_state');
  assert.match(textOf(result), /no Recorder/);
});
