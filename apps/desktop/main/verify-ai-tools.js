// AI Tools (owner, 2026-10-16: "it should be off unless some one opens it for
// single session or want to keep it always open"; 2026-10-17: "protect it from
// remote connections", "the ui should show indicate that and clicking should
// give a detail who is connected and from").
//
// The app menu's AI Tools opens the window with the tools off and no badge in
// the title bar. Always on asks for a folder first (cancelled, the tools stay
// off), then, with the folder chosen, puts the badge in the title bar. An AI app is then played here: Rutba Office's own program started
// as Node with the tools' entry (mcp/node.js), as the window says to start it,
// spoken to on standard input and output. It makes a document. The badge says
// one app is connected, and the window names it, says it is on this computer,
// and logs its call. Turned off, the next call is refused and the badge goes.
// Run alone with RUTBA_VERIFY_ONLY=ai-tools.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { app } from 'electron';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyAiTools(h) {
  const { open, check, until, wait, errorsIn } = h;
  const win = await open('home');
  const js = (code) => win.webContents.executeJavaScript(code);
  let child = null;
  try {
    // The app menu's AI Tools.
    await js(`window.dispatchEvent(new Event('rutba-ai-tools')), 1`);
    await until(() => js(`Boolean(document.querySelector('.ai-mode input:checked'))`), 'AI Tools and its setting', 6000).catch(() => {});
    const first = await js(`({ off: document.querySelector('.ai-mode input[value="off"]')?.checked, badge: Boolean(document.querySelector('.ai-badge')), local: document.querySelector('.ai-local')?.textContent || '' })`);
    check('ai tools: the app menu opens AI Tools with the tools off, no badge in the title bar, and the window saying no other machine can connect',
      first.off === true && first.badge === false && /Nothing listens on the network/.test(first.local), JSON.stringify(first));

    // Always on with no folder yet: the person chooses one first, and cancelling leaves the tools off.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [] });
    await js(`document.querySelector('.ai-mode input[value="always"]')?.click(), 1`);
    await wait(600);
    const cancelled = await js(`window.rutbaOffice.ai.state().then((s) => ({ mode: s.mode, folders: s.folders.length }))`);
    check('ai tools: turning them on with no folder asks for one first, and cancelling leaves them off with no folder made',
      cancelled.mode === 'off' && cancelled.folders === 0, JSON.stringify(cancelled));
    const chosen = path.join(app.getPath('userData'), 'Chosen for the AI');
    fs.mkdirSync(chosen, { recursive: true });
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [chosen] });
    await js(`document.querySelector('.ai-mode input[value="always"]')?.click(), 1`);
    const on = await until(async () => (await js(`document.querySelector('.ai-mode input[value="always"]')?.checked && Boolean(document.querySelector('.ai-badge'))`)) || null, 'the tools on', 6000).catch(() => false);
    const state = await js(`window.rutbaOffice.ai.state()`);
    const folder = state.folders?.[0] || '';
    check('ai tools: Always on, with the folder the person chose, puts the badge in the title bar',
      Boolean(on) && state.mode === 'always' && folder === chosen && state.folders.length === 1, `${state.mode}; ${folder}`);

    // An AI app, as the window says to start one.
    const { command, args, env } = state.connect;
    child = spawn(command, args, { stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, ...env, RUTBA_OFFICE_PROFILE: app.getPath('userData') } });
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
      const timer = setTimeout(() => resolve(null), 30000);
      waiting.set(my, (m) => { clearTimeout(timer); resolve(m); });
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`);
    });
    const init = await ask('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'Check client', version: '1.0' } });
    const made = await ask('tools/call', { name: 'create_document', arguments: { path: 'from-the-check.docx', title: 'Made by an AI app', blocks: [{ type: 'paragraph', text: 'Through the AI tools.' }] } });
    const file = path.join(folder, 'from-the-check.docx');
    check('ai tools: an AI app starts the tools as the window says, and they make a document in the folder',
      init?.result?.serverInfo?.name === 'rutba-office' && !made?.result?.isError && fs.existsSync(file), `${JSON.stringify(init?.result?.serverInfo)}; ${JSON.stringify(made?.result?.content?.[0]?.text || made)}`);

    // The badge counts it, and the window says who and from where.
    const badge = await until(async () => (await js(`document.querySelector('.ai-badge')?.dataset.connections || '0'`)) === '1' || null, 'the badge to count the app', 12000).then(() => true).catch(() => false);
    await js(`(() => { document.querySelector('.ai-tools-dialog .rw-btn.primary')?.click(); return 1; })()`);
    await wait(300);
    await js(`document.querySelector('.ai-badge')?.click(), 1`);
    const readShown = () => js(`(() => { const li = document.querySelector('.ai-connections li'); return li ? { who: li.querySelector('.ai-conn-who')?.textContent, where: li.querySelector('.ai-conn-where')?.textContent, last: li.querySelector('.ai-conn-last')?.textContent, log: document.querySelector('.ai-log li .ai-log-tool')?.textContent } : null; })()`);
    await until(readShown, 'who is connected', 6000).catch(() => null);
    const shown = await readShown();
    check('ai tools: while an AI app is connected the badge says so, and a click on it names the app, says it is on this computer, and shows its last call and the log',
      badge && shown?.who === 'Check client 1.0' && shown?.where === 'This computer' && /create_document/.test(shown?.last || '') && shown?.log === 'create_document',
      `badge ${badge}; ${JSON.stringify(shown)}`);

    // Off: the next call is refused, and the badge goes.
    await js(`document.querySelector('.ai-mode input[value="off"]')?.click(), 1`);
    await until(() => js(`window.rutbaOffice.ai.state().then((s) => s.mode === 'off')`), 'the tools off', 5000).catch(() => {});
    const refused = await ask('tools/call', { name: 'list_folders', arguments: {} });
    const badgeGone = await until(async () => !(await js(`Boolean(document.querySelector('.ai-badge'))`)) || null, 'the badge to go', 8000).then(() => true).catch(() => false);
    check('ai tools: turned off, the next call is refused with how to turn them on, and the badge goes',
      refused?.result?.isError === true && /AI tools are off/.test(refused?.result?.content?.[0]?.text || '') && badgeGone,
      `${JSON.stringify(refused?.result?.content?.[0]?.text || refused).slice(0, 160)}; badge gone ${badgeGone}`);

    child.stdin.end();
    await new Promise((r) => { child.once('exit', r); setTimeout(r, 5000); });
    child = null;
    await wait(300);
    const left = (await js(`window.rutbaOffice.ai.state()`)).connections?.length ?? -1;
    check('ai tools: an AI app that goes is no longer listed as connected', left === 0, `${left} listed`);

    const complaints = await errorsIn(win);
    check('ai tools: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('ai tools: the pass ran', false, err.stack || err.message);
  } finally {
    if (child) child.kill();
  }
}
