// What the AI tools may reach beyond files (owner, 2026-10-26, D-2026-10-26-04,
// -05, -06: "so the settings should be email accounts", "allow sending too,
// but confirm each message first with choice to create drafts only").
//
// A mail account is added (on a made-up server, never a real one). In Mail's
// Accounts its AI tools setting is Off; set to Read and prepare there, AI
// Tools shows it so; set to Read, prepare and act in AI Tools, Mail's
// Accounts shows that. Then an AI app is played as in verify-ai-tools.js
// (Rutba Office's own program run as Node, spoken to on standard input and
// output), with a test SMTP server started here on this computer: a message
// it asks to send is shown in the window with Send, Keep as draft and
// Cancel, and nothing reaches the server until Send is pressed; Keep as
// draft keeps a draft and sends nothing; Cancel does nothing. Set back to
// Read and prepare, a send is refused with where the setting is, nothing is
// shown, and a draft can still be made. Calendar's and Contacts' Accounts
// and Data Files' side panel show the same setting (verifyAiReachApps). Run
// alone with RUTBA_VERIFY_ONLY=ai-reach.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { app } from 'electron';
import { startTestSmtp } from './check-smtp.js';

/**
 * The same setting in Calendar's and Contacts' Accounts and in Data Files'
 * side panel: chosen there, the AI tools have it.
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyAiReachApps(h) {
  const { open, check, until, wait, errorsIn } = h;
  const choose = (js, selector, value) => js(`(() => { const s = document.querySelector(${JSON.stringify(selector)}); if (!s) return false; s.value = ${JSON.stringify(value)}; s.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
  for (const [which, kind, title] of [['calendar', 'calendar', 'Calendar\'s'], ['contacts', 'book', 'Contacts\'']]) {
    try {
      const win = await open(which);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Accounts'))`), 'Home, Accounts', 8000);
      await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Accounts').click(), 1`);
      const sel = `.ai-things-own[data-ai-kind="${kind}"] .ai-step-select`;
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), 'the AI tools settings in Accounts', 6000).catch(() => {});
      const id = await js(`document.querySelector(${JSON.stringify(sel)})?.dataset.aiId || null`);
      const before = await js(`document.querySelector(${JSON.stringify(sel)})?.value || null`);
      await choose(js, sel, 'read');
      await wait(500);
      const kept = id ? await js(`window.rutbaOffice.ai.stepOf({ kind: ${JSON.stringify(kind)}, id: ${JSON.stringify(id)} }).then((s) => s.step)`) : null;
      await js(`window.rutbaOffice.ai.setStep({ kind: ${JSON.stringify(kind)}, id: ${JSON.stringify(id)}, step: 'off' }).then(() => 1, () => 1)`);
      const complaints = await errorsIn(win);
      check(`ai reach: ${title} Accounts lists each ${kind === 'book' ? 'address book' : 'calendar'} with its AI tools setting, Off until chosen, and Read chosen there is what the tools have`,
        Boolean(id) && before === 'off' && kept === 'read' && complaints.length === 0, JSON.stringify({ id, before, kept, complaints }));
    } catch (err) {
      check(`ai reach: the ${which} pass ran`, false, err.stack || err.message);
    }
  }

  // Data Files: the side panel's AI tools, kept on this computer.
  const file = path.join(app.getPath('userData'), 'For the AI tools.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const { Base, buildFromSpec } = await import('@rutba/database');
    const b = Base.open(file);
    buildFromSpec(b, { tables: [{ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] }] });
    b.close();
    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-ai-tools'))`), 'the AI tools button', 8000).catch(() => {});
    const label = await js(`document.querySelector('.bs-ai-tools')?.textContent || ''`);
    await js(`document.querySelector('.bs-ai-tools')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-ai-dialog input[value="prepare"]'))`), 'the AI tools dialog', 5000).catch(() => {});
    await js(`document.querySelector('.bs-ai-dialog input[value="prepare"]')?.click(), 1`);
    await wait(200);
    await js(`document.querySelector('.bs-ai-save')?.click(), 1`);
    await until(async () => /Read and prepare/.test(await js(`document.querySelector('.bs-ai-tools')?.textContent || ''`)) || null, 'the button saying the step', 5000).catch(() => {});
    const after = await js(`document.querySelector('.bs-ai-tools')?.textContent || ''`);
    const listed = await js(`window.rutbaOffice.ai.things().then((r) => r.things.find((x) => x.kind === 'database' && x.id.toLowerCase() === ${JSON.stringify(file.toLowerCase())}) || null)`);
    const complaints = await errorsIn(win);
    check('ai reach: Data Files\' side panel says the database is Off for the AI tools until chosen, and Read and prepare chosen there is listed so in AI Tools',
      /AI tools: Off/.test(label) && /AI tools: Read and prepare/.test(after) && listed?.step === 'prepare' && complaints.length === 0,
      JSON.stringify({ label, after, listed, complaints }));
  } catch (err) {
    check('ai reach: the Data Files pass ran', false, err.stack || err.message);
  }
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyAiReach(h) {
  const { open, check, until, wait, errorsIn } = h;
  let child = null;
  let smtp = null;
  let accountId = null;
  let win = null;
  const js = (code) => win.webContents.executeJavaScript(code);
  try {
    win = await open('mail');
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the accounts', 8000);

    // The tools on, with a folder, as the person turns them on in AI Tools.
    const folder = path.join(app.getPath('userData'), 'For the AI reach check');
    fs.mkdirSync(folder, { recursive: true });
    await js(`window.rutbaOffice.ai.addFolder({ path: ${JSON.stringify(folder)} }).then(() => window.rutbaOffice.ai.setMode({ mode: 'always' })).then(() => 1)`);

    // An account on a made-up server: the check's transport hands what it sends to the test server below, and nothing else.
    accountId = (await js(`window.rutbaOffice.mail.addAccount({ account: { email: 'tools@checks.example', name: 'Tools Check', imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`)).id;
    win.webContents.reload();
    await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Tools Check/.test(b.textContent))`), 'the new account', 10000);

    const openAccounts = async () => {
      await js(`[...document.querySelectorAll('.ml-add-account button')].find((b) => /Manage accounts/.test(b.textContent)).click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.ml-account-row[data-account="tools@checks.example"] .ai-step-select'))`), 'the account\'s AI tools setting', 6000);
    };
    const closeDialog = () => js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => /^(Done|Close)$/.test(b.textContent.trim()))?.click(), 1`);
    const choose = (selector, value) => js(`(() => { const s = document.querySelector(${JSON.stringify(selector)}); if (!s) return false; s.value = ${JSON.stringify(value)}; s.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
    const inMail = '.ml-account-row[data-account="tools@checks.example"] .ai-step-select';
    const inTools = `.ai-tools-dialog .ai-thing[data-ai-kind="mail"][data-ai-id="${accountId}"] .ai-step-select`;
    const valueOf = (selector) => js(`document.querySelector(${JSON.stringify(selector)})?.value || null`);

    // Off until chosen; set in Mail's Accounts, shown in AI Tools.
    await openAccounts();
    const first = await js(`(() => { const s = document.querySelector(${JSON.stringify(inMail)}); return { value: s?.value, options: [...(s?.options || [])].map((o) => o.textContent) }; })()`);
    await choose(inMail, 'prepare');
    await until(async () => (await js(`window.rutbaOffice.ai.stepOf({ kind: 'mail', id: ${JSON.stringify(accountId)} }).then((s) => s.step)`)) === 'prepare' || null, 'the setting kept', 5000).catch(() => {});
    await closeDialog();
    await wait(300);
    await js(`window.dispatchEvent(new Event('rutba-ai-tools')), 1`);
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(inTools)}))`), 'the account in AI Tools', 6000).catch(() => {});
    const shownInTools = await valueOf(inTools);
    check('ai reach: a mail account\'s AI tools setting is Off until chosen, offers the four steps in Mail\'s Accounts, and Read and prepare set there shows in AI Tools',
      first.value === 'off' && first.options.join('|') === 'Off|Read|Read and prepare|Read, prepare and act' && shownInTools === 'prepare',
      JSON.stringify({ first, shownInTools }));

    // Set in AI Tools, shown in Mail's Accounts.
    await choose(inTools, 'act');
    await until(async () => (await valueOf(inTools)) === 'act' || null, 'the setting in AI Tools', 5000).catch(() => {});
    await closeDialog();
    await wait(300);
    await openAccounts();
    const backInMail = await valueOf(inMail);
    await closeDialog();
    check('ai reach: Read, prepare and act set in AI Tools shows in Mail\'s Accounts, the one setting', backInMail === 'act', String(backInMail));

    // An AI app, as AI Tools says to start one, and a test SMTP server on this computer.
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
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
      const timer = setTimeout(() => resolve(null), 60000);
      waiting.set(my, (m) => { clearTimeout(timer); resolve(m); });
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`);
    });
    const textOf = (m) => m?.result?.content?.map((c) => c.text).join('\n') || JSON.stringify(m?.error || m);
    await ask('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'Reach check', version: '1.0' } });
    const send = (subject) => ask('tools/call', { name: 'send_message', arguments: { account: 'tools@checks.example', to: 'friend@checks.example', subject, text: 'Sent through the AI tools, only on the press.' } });
    const shownConfirm = async () => {
      const there = await until(() => js(`Boolean(document.querySelector('.ai-confirm .ai-confirm-act'))`), 'the confirmation', 15000).then(() => true, () => false);
      if (!there) return null;
      return js(`(() => { const d = document.querySelector('.ai-confirm'); return { to: d.querySelector('.ai-confirm-to')?.textContent || '', subject: d.querySelector('.ai-confirm-subject')?.textContent || '', act: d.querySelector('.ai-confirm-act')?.textContent.trim() || '', draft: Boolean(d.querySelector('.ai-confirm-draft')), cancel: Boolean(d.querySelector('.ai-confirm-cancel')) }; })()`);
    };
    const press = (which) => js(`document.querySelector('.ai-confirm .ai-confirm-${which}')?.click(), 1`);
    const drafts = async () => (await js(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Drafts', offset: 0, limit: 50 }).then((r) => r.rows.map((m) => m.subject), () => [])`)) || [];

    // Send: shown first, and only the press sends it.
    const sending = send('Only on the press');
    const shown = await shownConfirm();
    await wait(1500);
    const beforePress = smtp.received.length;
    await press('act');
    const sent = await sending;
    await until(() => smtp.received.length > 0 || null, 'the message at the test server', 8000).catch(() => {});
    const got = smtp.received[0];
    check('ai reach: a message an AI app asks to send is shown in Rutba Office with To, Subject, Send, Keep as draft and Cancel, and reaches the test server only when Send is pressed',
      shown?.to === 'friend@checks.example' && shown?.subject === 'Only on the press' && shown?.act === 'Send' && shown?.draft && shown?.cancel
        && beforePress === 0 && smtp.received.length === 1 && got?.to?.[0] === 'friend@checks.example' && /Only on the press/.test(got?.data || '') && /Sent from tools@checks\.example/.test(textOf(sent)),
      JSON.stringify({ shown, beforePress, received: smtp.received.length, to: got?.to, said: textOf(sent).slice(0, 160) }));

    // Keep as draft: a draft, nothing sent.
    const keeping = send('Kept, not sent');
    const shownAgain = await shownConfirm();
    await press('draft');
    const kept = await keeping;
    await wait(500);
    const keptDrafts = await drafts();
    check('ai reach: Keep as draft keeps the message as a draft in the account and sends nothing',
      Boolean(shownAgain) && /Kept as a draft/.test(textOf(kept)) && smtp.received.length === 1 && keptDrafts.includes('Kept, not sent'),
      JSON.stringify({ said: textOf(kept).slice(0, 160), received: smtp.received.length, drafts: keptDrafts }));

    // Cancel: nothing.
    const cancelling = send('Cancelled');
    const shownThird = await shownConfirm();
    await press('cancel');
    const cancelled = await cancelling;
    await wait(500);
    const afterCancel = await drafts();
    const gone = await js(`!document.querySelector('.ai-confirm')`);
    check('ai reach: Cancel sends nothing and keeps nothing, and the confirmation goes',
      Boolean(shownThird) && /Cancelled/.test(textOf(cancelled)) && smtp.received.length === 1 && !afterCancel.includes('Cancelled') && gone,
      JSON.stringify({ said: textOf(cancelled).slice(0, 160), received: smtp.received.length, drafts: afterCancel, gone }));

    // Drafts only: a send is refused with where the setting is; nothing is shown; a draft can still be made.
    await js(`window.rutbaOffice.ai.setStep({ kind: 'mail', id: ${JSON.stringify(accountId)}, step: 'prepare' }).then(() => 1)`);
    const refused = await send('Not allowed');
    await wait(1000);
    const noConfirm = await js(`!document.querySelector('.ai-confirm')`);
    const drafted = await ask('tools/call', { name: 'create_draft', arguments: { account: 'tools@checks.example', to: 'friend@checks.example', subject: 'A draft only', text: 'x' } });
    await wait(300);
    const draftsNow = await drafts();
    check('ai reach: set to Read and prepare, a send is refused naming the account and where its setting is, nothing is shown or sent, and a draft is still made',
      refused?.result?.isError === true && /tools@checks\.example: it is set to Read and prepare, and this needs Read, prepare and act\. Its setting is in Mail under Accounts, or in AI Tools/.test(textOf(refused))
        && noConfirm && smtp.received.length === 1 && !drafted?.result?.isError && draftsNow.includes('A draft only'),
      JSON.stringify({ refused: textOf(refused).slice(0, 220), noConfirm, received: smtp.received.length, drafts: draftsNow }));

    // The log: the tool, the account and the step; never what was in the message.
    const log = await js(`window.rutbaOffice.ai.state().then((s) => s.log.slice(0, 6))`);
    check('ai reach: the log names the tool, the account and the step, and never the message',
      log.some((e) => e.tool === 'send_message' && e.thing === 'tools@checks.example' && e.step === 'act') && log.some((e) => e.tool === 'create_draft' && e.step === 'prepare') && !JSON.stringify(log).includes('only on the press'),
      JSON.stringify(log).slice(0, 400));

    const complaints = await errorsIn(win);
    check('ai reach: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('ai reach: the pass ran', false, err.stack || err.message);
  } finally {
    if (child) {
      child.stdin.end();
      await new Promise((r) => { child.once('exit', r); setTimeout(r, 5000); });
    }
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close();
    if (win && !win.isDestroyed()) {
      await js(`window.rutbaOffice.ai.setMode({ mode: 'off' }).then(() => 1, () => 1)`).catch(() => {});
      if (accountId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1, () => 1)`).catch(() => {});
    }
  }
}
