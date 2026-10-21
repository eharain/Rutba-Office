// An automated accessibility audit of every app, as a screen reader or a
// keyboard user meets it: axe-core (MPL-2.0, a development dependency the
// checks alone load; nothing of it ships) run over each app's window as it
// opens on real content, over each ribbon tab in turn, and over the menus,
// dialogs and panes a person opens most. A serious or critical violation
// fails the check and is listed by rule and element; moderate and minor ones
// are printed for the record. The page's own content stands apart where the
// app does not draw it: a message's HTML (the sender's) and a PDF's own
// text layer are left to their own checks (screen-reader). Run alone with
// RUTBA_VERIFY_ONLY=a11y-audit; captures a11y-<app>.png.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fixtures, message } from './verify-polish-pass.js';

const require = createRequire(import.meta.url);
let axeSource = null;
/** axe-core's own script, read once from the development dependency. */
export function axeScript() {
  if (axeSource === null) axeSource = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
  return axeSource;
}

/** What the audit leaves to other checks: content the app shows but did not write. */
const EXCLUDE = [['iframe[srcdoc]'], ['.textLayer'], ['.annotationLayer']];

/**
 * axe over the page, as JSON a check can print: each violation with its
 * impact, its help text and the elements it found, by selector.
 */
export async function runAxe(win, { include = null } = {}) {
  const js = (code) => win.webContents.executeJavaScript(code);
  const loaded = await js(`typeof window.axe === 'object'`);
  if (!loaded) await js(`${axeScript()};\n1`);
  const context = include ? `{ include: ${JSON.stringify(include)}, exclude: ${JSON.stringify(EXCLUDE)} }` : `{ exclude: ${JSON.stringify(EXCLUDE)} }`;
  return js(`(async () => {
    const r = await window.axe.run(${context}, {
      runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice'] },
      resultTypes: ['violations'],
      rules: { 'color-contrast': { enabled: true }, region: { enabled: true } },
    });
    return r.violations.map((v) => ({
      id: v.id, impact: v.impact, help: v.help,
      nodes: v.nodes.slice(0, 12).map((n) => ({ target: n.target.join(' '), html: n.html.slice(0, 160), why: (n.failureSummary || '').split('\\n').slice(1, 3).join(' ').trim().slice(0, 200) })),
      count: v.nodes.length,
    }));
  })()`);
}

const SERIOUS = new Set(['serious', 'critical']);

/**
 * Findings WCAG itself sets aside, each with the reason, printed as noted
 * rather than failed. axe cannot see an equivalent control elsewhere.
 */
export const EXCUSED = [
  { id: 'target-size', at: /\.ml-check|\.ml-star/, why: "WCAG 2.5.8's Equivalent exception: a message is checked by Ctrl+click or Shift+click anywhere on its row, the whole row's height, and starred from the ribbon's Star and the row's menu" },
  { id: 'scrollable-region-focusable', at: /^\.sh-grid$/, why: 'the grid scroller sits inside the sheet (.sh), which is in the Tab order and scrolls it with the arrow keys, Page Up and Page Down' },
];
/**
 * Known not to meet a criterion, and said so in docs/ACCESSIBILITY.md: noted
 * on every run, never silently passed. Taking it off this list is the work.
 */
export const KNOWN = [
  { id: 'target-size', at: /\.cal-chip|class="cal-chip/, why: "the month view's events are 16 px tall so a busy day shows its events at all; the same events are full-size targets in the Day, Week and Agenda views and in the day's +N more list" },
];
const excused = (id, target) => EXCUSED.find((e) => e.id === id && e.at.test(target));
const known = (id, target, html = '') => KNOWN.find((e) => e.id === id && (e.at.test(target) || e.at.test(html)));

export async function verifyA11yAudit({ open, check, until, wait, press }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const only = (process.env.RUTBA_A11Y_APPS || '').split(',').map((x) => x.trim()).filter(Boolean);
  const wanted = (app) => !only.length || only.includes(app);
  const fx = fixtures(dir, 'a11y-fixtures');
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };

  /** One app: its window, each ribbon tab, then what `extra` opens; every serious finding named. */
  const audit = async (app, win, extra = []) => {
    const js = (code) => win.webContents.executeJavaScript(code);
    const found = new Map();
    const minor = new Map();
    const noted = new Map();
    const take = (where, list) => {
      for (const v of list) {
        for (const n of v.nodes) {
          const key = `${v.id}|${n.target}`;
          const into = !SERIOUS.has(v.impact) || excused(v.id, n.target) ? minor : known(v.id, n.target, n.html) ? noted : found;
          if (!into.has(key)) into.set(key, { where, id: v.id, impact: v.impact, help: v.help, target: n.target, html: n.html, why: n.why });
        }
      }
    };
    take('window', await runAxe(win));
    await snap(win, `a11y-${app}.png`);
    const tabs = await js(`[...document.querySelectorAll('.rw-tabs [role="tab"]')].map((t) => t.textContent.trim())`);
    for (const [i, name] of tabs.entries()) {
      await js(`document.querySelectorAll('.rw-tabs [role="tab"]')[${i}]?.click(), 1`);
      await until(() => js(`document.querySelectorAll('.rw-tabs [role="tab"]')[${i}]?.getAttribute('aria-selected') === 'true'`), `the ${name} tab`, 4000).catch(() => {});
      await wait(250);
      take(`${name} tab`, await runAxe(win, { include: [['.rw-ribbon']] }));
    }
    if (tabs.length) await js(`document.querySelectorAll('.rw-tabs [role="tab"]')[0]?.click(), 1`);
    for (const step of extra) {
      try {
        const shown = await step.open(js, win);
        if (shown === false) { check(`a11y-audit: ${app}: ${step.name} opens for the audit`, false, 'it never opened'); continue; }
        await wait(300);
        take(step.name, await runAxe(win, step.include ? { include: step.include } : {}));
        if (step.shot) await snap(win, `a11y-${app}-${step.shot}.png`);
        if (step.close) await step.close(js);
      } catch (err) {
        check(`a11y-audit: ${app}: ${step.name} audited`, false, err.message);
      }
    }
    const faults = [...found.values()];
    for (const f of faults) console.log(`     [a11y-audit] ${app} (${f.where}) ${f.impact} ${f.id}: ${f.target} :: ${f.html.replace(/\s+/g, ' ')} :: ${f.why}`);
    for (const f of minor.values()) console.log(`     [a11y-audit] ${app} (${f.where}) ${f.impact} ${f.id}: ${f.target}`);
    for (const f of noted.values()) console.log(`     [a11y-audit] ${app} (${f.where}) known, not met: ${f.id}: ${f.target} (${known(f.id, f.target, f.html).why})`);
    check(`a11y-audit: ${app}: its window, ${tabs.length} ribbon tab${tabs.length === 1 ? '' : 's'}${extra.length ? ` and ${extra.map((s) => s.name).join(', ')}` : ''} have no serious or critical violation (axe-core, WCAG 2.2 A and AA)`,
      faults.length === 0, faults.length ? `${faults.length}: ` + faults.slice(0, 8).map((f) => `${f.id} at ${f.target} (${f.where})`).join('; ') : `${minor.size} moderate or minor noted${noted.size ? `; ${noted.size} known not to meet ${[...new Set([...noted.values()].map((f) => f.id))].join(', ')} (docs/ACCESSIBILITY.md)` : ''}`);
  };

  const ready = (win, sel, what) => until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, 15000).catch(() => {});
  const escape = async (js, gone) => {
    await js(`document.activeElement?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    await until(() => js(gone), 'it to close', 4000).catch(() => {});
  };
  const opener = (selector, waitFor) => async (js) => {
    const pressed = await js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); b?.click(); return Boolean(b); })()`);
    if (!pressed) return false;
    return until(() => js(`Boolean(document.querySelector(${JSON.stringify(waitFor)}))`), waitFor, 5000).then(() => true).catch(() => false);
  };
  const appMenu = { name: 'the app menu', open: opener('.rw-appmark', '.rw-menu'), include: [['.rw-menu']], close: (js) => escape(js, `!document.querySelector('.rw-menu')`) };
  const dialogBy = (name, selector, shot) => ({ name, shot, open: opener(selector, '.rw-dialog'), include: [['.rw-dialog']], close: (js) => escape(js, `!document.querySelector('.rw-dialog')`) });
  // A dialog, bar or pane the app's own shortcut opens, audited with the whole window around it; left open
  // at the end of the app's turn when `last` (a form whose Escape would ask a question), else closed.
  const byKey = (name, key, waitFor, shot, last = false) => ({
    name, shot,
    open: async (js, win) => {
      if (!press) return false;
      await js(`(document.querySelector('.rw-body') || document.body).focus?.(), 1`);
      await press(win.webContents, key, { modifiers: ['control'] });
      return until(() => js(`Boolean(document.querySelector(${JSON.stringify(waitFor)}))`), waitFor, 5000).then(() => true).catch(() => false);
    },
    close: last ? null : (js) => escape(js, `!document.querySelector(${JSON.stringify(waitFor)})`),
  });

  const apps = [
    ['home', undefined, '.home-card', [appMenu]],
    ['word', fx.docx, '.wd-page', [appMenu, dialogBy('Print', '[data-tip^="Print"]', 'print'), byKey('Find and Replace', 'F', '.rw-dialog, .wd-find', 'find')]],
    ['sheets', fx.xlsx, '.sh-cell', [appMenu, byKey('Insert Link', 'K', '.rw-dialog', 'link')]],
    ['slides', fx.pptx, '.sl-thumb', [appMenu, byKey('Replace', 'H', '.rw-dialog, .sl-find', 'replace')]],
    ['pdf', fx.pdf, '.pdf-page canvas[data-drawn]', [appMenu, byKey('Find', 'F', '.pdf-find', 'find')]],
    ['pictures', fx.pics, '.rw-ribbon .rw-tab', [appMenu]],
    ['image', fx.png, '.rw-ribbon .rw-tab', [appMenu]],
    ['video', fx.png, '.vd-clip', [appMenu]],
    ['contacts', fx.vcf, '.ct-item', [appMenu, byKey('a new contact', 'N', '.ct-editor', 'new', true)]],
    ['calendar', fx.ics, '.cal-month-grid, .cal-time', [appMenu, byKey('a new event', 'N', '.rw-dialog', 'event', true)]],
  ];
  for (const [app, file, sel, extra] of apps.filter(([a]) => wanted(a))) {
    try {
      const win = await open(app, file);
      await ready(win, sel, `the ${app} window`);
      await wait(500);
      await audit(app, win, extra);
    } catch (err) {
      check(`a11y-audit: the ${app} audit ran`, false, err.stack || err.message);
    }
  }

  if (wanted('mail')) {
    let accountId = null;
    let js = null;
    try {
      const win = await open('mail');
      js = (code) => win.webContents.executeJavaScript(code);
      const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
      await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0 || Boolean(document.querySelector('.rw-empty'))`), 'the Mail window', 8000).catch(() => {});
      const made = await mail('addAccount', { account: { email: 'a11y@checks.example', name: 'Audit (checks)', local: true } });
      accountId = made.id;
      const now = Date.now();
      await mail('deliverTest', { accountId, folder: 'Inbox', raw: message({ from: 'Sara Khan <sara.khan@example.com>', subject: 'The spring range', body: 'The samples arrive on Tuesday.', date: new Date(now - 3600e3).toUTCString() }) });
      await mail('deliverTest', { accountId, folder: 'Inbox', raw: message({ from: 'Lena Weiss <lena.weiss@example.com>', subject: 'The autumn letter', body: 'Three things happened this month.', date: new Date(now).toUTCString() }) });
      await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => /Audit \\(checks\\)|a11y@checks/.test(n.textContent)); b?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.ml-row').length >= 2`), 'the two messages', 8000).catch(() => {});
      await js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((n) => /autumn letter/.test(n.textContent)); r?.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.ml-reader'))`), 'the message open', 6000).catch(() => {});
      await wait(500);
      // The equivalent the target-size exception leans on: the whole row checks a message.
      const rowClick = (i, mods) => js(`(() => { const r = document.querySelectorAll('.ml-row')[${i}]; r?.dispatchEvent(new MouseEvent('click', { bubbles: true, ${mods} })); return Boolean(r); })()`);
      const checkedNow = () => js(`document.querySelectorAll('.ml-list > .ml-row.checked').length`);
      await rowClick(0, 'ctrlKey: true');
      const one = await until(() => checkedNow().then((n) => n === 1), 'one row checked', 3000).then(() => true).catch(() => false);
      await rowClick(1, 'shiftKey: true');
      const two = await until(() => checkedNow().then((n) => n === 2), 'two rows checked', 3000).then(() => true).catch(() => false);
      const stillOpen = await js(`/autumn letter/.test(document.querySelector('.ml-reader')?.textContent || '')`);
      check('a11y-audit: mail: Ctrl+click on a row checks that message and Shift+click checks the run to another, without opening it, so the small check box has an equivalent the size of the row',
        one && two && stillOpen, JSON.stringify({ one, two, stillOpen }));
      await js(`document.querySelector('.ml-bulk .ml-check')?.click(), 1`);
      await audit('mail', win, [appMenu, byKey('a new message', 'N', '.ml-compose', 'compose', true)]);
    } catch (err) {
      check('a11y-audit: the mail audit ran', false, err.stack || err.message);
    } finally {
      if (accountId && js) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} })`).catch(() => {});
    }
  }
}
