// The Recorder's AI tools (2.22, D-2026-10-30-06): an AI app, played as in
// verify-ai-reach.js (Rutba Office's own program run as Node, spoken to on
// standard input and output), asks the Recorder for its windows while the
// Recorder is Off for the tools, and is refused; set to record in AI Tools,
// it lists the windows, records a Word window for five seconds into the
// folder the tools were given, and stops: the file it is told of is there,
// plays, and is as long as the answer said. The Recorder's own settings are
// as the person left them afterwards. Run alone with
// RUTBA_VERIFY_ONLY=recorder-tools.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { app, BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { seedSetup, readSetup } from './verify-studio-seed.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyRecorderTools(h, { dir }) {
  const { open, check, wait } = h;
  const standIn = globalThis.__rutbaScreenStandIn;
  globalThis.__rutbaScreenStandIn = false;
  let child = null;
  let js = null;
  const folder = path.join(app.getPath('userData'), 'For the recorder tools check');
  try {
    const docFile = path.join(dir, 'rec-tools-window.docx');
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded through the AI tools.' }] }));
    const win = await open('word', docFile);
    js = (code) => win.webContents.executeJavaScript(code);
    fs.rmSync(folder, { recursive: true, force: true });
    fs.mkdirSync(folder, { recursive: true });
    await js(`window.rutbaOffice.ai.addFolder({ path: ${JSON.stringify(folder)} }).then(() => window.rutbaOffice.ai.setMode({ mode: 'always' })).then(() => 1)`);
    // The person's own Recorder settings, unlike what the tools ask for in every way that matters,
    // read by a Recorder the tools open: one an earlier check left open holds the settings it read then.
    for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed() && /app=recorder\b/.test(w.webContents.getURL())) w.destroy();
    await wait(300);
    const mine = { size: 720, fps: 24, countdown: 5, folder: path.join(dir, 'my recordings') };
    // The person's share: their camera alone, as a circle, with their title over it, and no sound from the computer.
    const myStack = { version: 2, video: 'on', sound: 'on', scenes: [],
      items: [{ id: 'title-1', kind: 'title', text: 'My own title', on: true, live: true, box: { x: 0.04, y: 0.78, w: 0.5, h: 0.14 } }, { id: 'camera-1', kind: 'camera', ref: '', on: true, live: true, shape: 'circle', box: { x: 0.3, y: 0.1, w: 0.4, h: 0.71 } }],
      sounds: [{ id: 'mic-1', kind: 'mic', ref: '', live: true, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', live: false, muted: false, gain: 1 }] };
    await seedSetup(js, { app: 'recorder', stack: myStack, studio: mine });
    const whatIsKept = async () => { const [doc, global] = await Promise.all([readSetup(js), js(`window.rutbaOffice.store.get({ key: 'studio.global', fallback: null })`)]); return { doc, settings: { ...doc?.settings, ...global } }; };
    const was = await whatIsKept();
    const settingsWere = was.settings;
    const stackWas = was.doc?.stack;
    const listedOff = await js(`window.rutbaOffice.ai.things().then((r) => r.things.find((x) => x.kind === 'recorder') || null)`);

    // An AI app, as AI Tools says to start one.
    const { command, args, env } = (await js(`window.rutbaOffice.ai.state()`)).connect;
    child = spawn(command, args, { stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, ...env } });
    let out = '';
    const waiting = new Map();
    child.stdout.on('data', (d) => {
      out += d;
      let at;
      while ((at = out.indexOf('\n')) >= 0) {
        const line = out.slice(0, at);
        out = out.slice(at + 1);
        try { const m = JSON.parse(line); waiting.get(m.id)?.(m); } catch { /* not ours */ }
      }
    });
    let id = 0;
    const ask = (method, params) => new Promise((resolve) => {
      const my = ++id;
      const timer = setTimeout(() => resolve(null), 90000);
      waiting.set(my, (m) => { clearTimeout(timer); resolve(m); });
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`);
    });
    const textOf = (m) => m?.result?.content?.map((c) => c.text).join('\n') || JSON.stringify(m?.error || m);
    const tool = (name, a = {}) => ask('tools/call', { name, arguments: a });
    await ask('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'Recorder tools check', version: '1.0' } });
    const offered = ((await ask('tools/list', {}))?.result?.tools || []).map((t) => t.name).filter((n) => n.startsWith('recorder_'));

    // Off: refused, naming where the setting is, and no Recorder opened for it.
    const refused = await tool('recorder_windows');
    check('recorder tools: the Recorder is Off for the AI tools until chosen, and a call then is refused with where its setting is',
      listedOff?.step === 'off' && offered.length === 4 && refused?.result?.isError === true && /set to Off.*AI Tools, under Recorder/.test(textOf(refused)),
      JSON.stringify({ listedOff, offered, refused: textOf(refused) }));

    // Set to record in AI Tools: the windows, a recording of the Word window, five seconds, stopped.
    await js(`window.rutbaOffice.ai.setStep({ kind: 'recorder', id: 'recorder', step: 'prepare' }).then(() => 1)`);
    const listed = textOf(await tool('recorder_windows'));
    const row = listed.split('\n').find((l) => l.includes('rec-tools-window'));
    const source = row ? row.split('\t')[1] : null;
    const started = source ? await tool('recorder_start', { source, folder, title: 'Recorded by the tools' }) : null;
    const during = textOf(await tool('recorder_state'));
    await wait(5000);
    const stopped = await tool('recorder_stop');
    const said = /Kept (.+): (\d+) seconds, (\d+) bytes\./.exec(textOf(stopped));
    const file = said?.[1] || null;
    const there = file && fs.existsSync(file) && path.dirname(file).toLowerCase() === folder.toLowerCase();
    const probe = there ? probeMedia(new Uint8Array(fs.readFileSync(file)), { complete: true }) : null;
    const after = textOf(await tool('recorder_state'));
    await wait(500);
    const now = await whatIsKept();
    const settingsNow = now.settings;
    const stackNow = now.doc?.stack;
    const savedAside = (await js(`window.rutbaOffice.store.get({ key: 'studio.setups', fallback: null })`))?.setups?.some((x) => x.tools);
    check('recorder tools: set to record, an AI app lists the windows, records one for five seconds into its folder and stops; the file it is told of plays, as long as it was told',
      Boolean(source) && started && !started.result?.isError && /^phase\trecording/.test(during) && there && probe?.hasVideo && probe.duration >= 4 && probe.duration <= 8
        && Math.abs(Number(said[2]) - probe.duration) <= 1.5 && Number(said[3]) === fs.statSync(file).size && /^phase\tsetup/.test(after),
      JSON.stringify({ source, started: textOf(started), during, stopped: textOf(stopped), probe: probe ? { video: probe.hasVideo, duration: probe.duration } : null, after }));
    const keys = Object.keys(mine);
    const pick = (o) => JSON.stringify(Object.fromEntries(keys.map((k) => [k, o?.[k]])));
    const shareOf = (st) => JSON.stringify((st?.items || []).map((it) => [it.id, it.kind, it.on, it.shape, it.text || '']));
    check('recorder tools: the Recorder\'s own settings and share are as the person left them once the tools\' recording ends',
      pick(settingsNow) === pick(settingsWere) && settingsWere?.countdown === 5 && shareOf(stackNow) === shareOf(stackWas) && (stackWas?.items || []).length === 2 && !savedAside,
      JSON.stringify({ settingsWere, settingsNow, stackWas: shareOf(stackWas), stackNow: shareOf(stackNow) }).slice(0, 900));

    // The log: the tools and the Recorder at the step they used.
    const log = await js(`window.rutbaOffice.ai.state().then((s) => s.log.slice(0, 8))`);
    check('recorder tools: the log names each Recorder tool, the Recorder and the step',
      ['recorder_start', 'recorder_stop'].every((n) => log.some((e) => e.tool === n && e.thing === 'Recorder' && e.step === 'prepare' && e.ok)),
      JSON.stringify(log).slice(0, 500));
  } catch (err) {
    check('recorder tools: the pass ran', false, err.stack || err.message);
  } finally {
    globalThis.__rutbaScreenStandIn = standIn;
    if (child) {
      child.stdin.end();
      await new Promise((r) => { child.once('exit', r); setTimeout(r, 5000); });
    }
    // As it found them: the Recorder Off, the tools off, and no folder of this check's left given.
    if (js) await js(`window.rutbaOffice.ai.setStep({ kind: 'recorder', id: 'recorder', step: 'off' }).then(() => window.rutbaOffice.ai.setMode({ mode: 'off' })).then(() => window.rutbaOffice.ai.removeFolder({ path: ${JSON.stringify(folder)} })).then(() => 1, () => 1)`).catch(() => {});
  }
}
