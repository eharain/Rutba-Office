// The dark theme over every app, as a person who chose it sees it: the
// choice kept the way the app menu keeps it and each window opened on it,
// so what the windows draw from the theme in code follows too, not only
// their style sheets. Each app shows real content (a document with headings,
// a sheet, a deck, a PDF, Mail with an HTML message open, Calendar's month
// and week with a busy day, a contact's card, pictures, a picture being
// edited, a video's clips and Home) at 1300x700 and 1680x1050, captured as
// dark-<app>-<size>.png in RUTBA_VERIFY_CAPTURE. Every piece of text on
// screen is measured against what is drawn under it: 4.5:1 for body text,
// 3:1 for large text (WCAG 2), disabled controls aside, as WCAG sets them
// aside. The choice is put back to the checks' light when done, and the
// Mail account is the pass's own and goes with it. Run alone with
// RUTBA_VERIFY_ONLY=dark-pass.

import fs from 'node:fs';
import path from 'node:path';
import { fixtures, message } from './verify-polish-pass.js';

const SIZES = [[1300, 700], [1680, 1050]];

/**
 * Text on screen that does not read: each element with words of its own,
 * shown and on top where its words are, with the colour of its words
 * against the colours under it composed down to the first solid one.
 */
const AUDIT_FN = `((w) => {
  const document = w.document, innerWidth = w.innerWidth, innerHeight = w.innerHeight;
  const getComputedStyle = (n) => w.getComputedStyle(n);
  const parse = (c) => {
    if (!c) return null;
    let m = /^rgba?\\(([^)]+)\\)$/.exec(c);
    if (m) { const p = m[1].split(/[\\s,\\/]+/).filter(Boolean).map(Number); return [p[0], p[1], p[2], p.length > 3 ? p[3] : 1]; }
    m = /^color\\(srgb ([^)]+)\\)$/.exec(c);
    if (m) { const p = m[1].split(/[\\s\\/]+/).filter(Boolean).map(Number); return [p[0] * 255, p[1] * 255, p[2] * 255, p.length > 3 ? p[3] : 1]; }
    return null;
  };
  const over = (top, under) => { const a = top[3]; return [top[0] * a + under[0] * (1 - a), top[1] * a + under[1] * (1 - a), top[2] * a + under[2] * (1 - a), 1]; };
  const lum = ([r, g, b]) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }; return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b); };
  const ratio = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
  const hex = (c) => '#' + c.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
  const page = parse(getComputedStyle(document.body).backgroundColor) || [255, 255, 255, 1];
  const under = (el) => {
    const layers = [];
    for (let n = el; n && n.nodeType === 1; n = n.parentElement) {
      const cs = getComputedStyle(n);
      if (cs.backgroundImage && cs.backgroundImage !== 'none' && !/^url/.test(cs.backgroundImage)) return null;
      const bg = parse(cs.backgroundColor);
      if (bg && bg[3] > 0) { layers.push(bg); if (bg[3] >= 0.99) break; }
    }
    let c = layers.length && layers[layers.length - 1][3] >= 0.99 ? layers.pop() : (page[3] >= 0.99 ? page : [255, 255, 255, 1]);
    while (layers.length) c = over(layers.pop(), c);
    return c;
  };
  const seen = new Set();
  const bad = [];
  let measured = 0;
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (!t.data.trim()) continue;
    const el = t.parentElement;
    if (!el || seen.has(el)) continue;
    seen.add(el);
    if (el.closest('svg, canvas, .textLayer, [disabled], [aria-disabled="true"], .disabled, [aria-hidden="true"], option, script, style')) continue;
    const range = document.createRange();
    range.selectNodeContents(t);
    const r = range.getBoundingClientRect();
    if (r.width < 2 || r.height < 2 || r.right < 0 || r.bottom < 0 || r.left > innerWidth || r.top > innerHeight) continue;
    const cx = Math.min(innerWidth - 1, Math.max(0, r.left + r.width / 2)), cy = Math.min(innerHeight - 1, Math.max(0, r.top + r.height / 2));
    const hit = document.elementFromPoint(cx, cy);
    if (!hit || !(hit === el || el.contains(hit) || hit.contains(el))) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility !== 'visible') continue;
    let opacity = 1;
    for (let n = el; n && n.nodeType === 1; n = n.parentElement) opacity *= Number(getComputedStyle(n).opacity);
    // Faded to show a control is off: set aside, as WCAG sets inactive controls aside.
    if (opacity < 0.9) continue;
    const fg = parse(cs.color);
    if (!fg || fg[3] === 0) continue;
    const bg = under(el);
    if (!bg) continue;
    const ink = fg[3] < 1 ? over(fg, bg) : fg;
    const size = parseFloat(cs.fontSize);
    const large = size >= 24 || (size >= 18.6 && Number(cs.fontWeight) >= 700);
    const need = large ? 3 : 4.5;
    const got = ratio(ink, bg);
    measured++;
    if (got < need - 0.05) {
      const where = (n) => (n.className && typeof n.className === 'string' ? '.' + n.className.trim().split(/\\s+/).slice(0, 2).join('.') : n.tagName.toLowerCase());
      bad.push({ at: where(el.parentElement || el) + ' > ' + where(el), text: t.data.trim().slice(0, 32), ink: hex(ink), bg: hex(bg), ratio: Math.round(got * 100) / 100, need });
    }
  }
  bad.sort((a, b) => a.ratio - b.ratio);
  return { measured, bad };
})`;
export const AUDIT = `${AUDIT_FN}(window)`;
/**
 * The same over each message frame: its document drawn again in a frame
 * without the sandbox (which lets no script in, ours included), the same
 * size, out of sight, measured from the window and taken away.
 */
const FRAMES_AUDIT = `(async () => {
  const out = { measured: 0, bad: [] };
  for (const f of [...document.querySelectorAll("iframe[srcdoc]")]) {
    const r = f.getBoundingClientRect();
    if (r.width < 10 || r.height < 10) continue;
    const copy = document.createElement("iframe");
    copy.style.cssText = "position: fixed; left: -30000px; top: 0; border: 0; width: " + r.width + "px; height: " + r.height + "px";
    const loaded = new Promise((done) => copy.addEventListener("load", done, { once: true }));
    copy.srcdoc = f.srcdoc;
    document.body.append(copy);
    await loaded;
    const got = (${AUDIT_FN})(copy.contentWindow);
    copy.remove();
    out.measured += got.measured;
    out.bad.push(...got.bad);
  }
  return out;
})()`;

export async function verifyDarkPass({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const only = (process.env.RUTBA_DARK_APPS || '').split(',').map((x) => x.trim()).filter(Boolean);
  const wanted = (app) => !only.length || only.includes(app);
  const fx = fixtures(dir, 'dark-pass-fixtures');
  let keep = null;
  let accountId = null;
  let mailJs = null;
  const report = {};
  try {
    // The choice, kept as the app menu keeps it: every window opened after it is dark.
    const first = await open('home');
    keep = (code) => first.webContents.executeJavaScript(code);
    await keep(`window.rutbaOffice.store.set({ key: 'theme', value: 'dark' }).then(() => 1)`);

    const look = async (app, win, setup = null) => {
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.documentElement.dataset.theme === 'dark'`), `the ${app} window dark`, 8000);
      if (setup) await setup(js);
      const found = [];
      for (const [w, h] of SIZES) {
        win.setContentSize(w, h);
        await until(() => js('[innerWidth, innerHeight]').then(([iw, ih]) => Math.abs(iw - w) <= 2 && Math.abs(ih - h) <= 2), `the window at ${w}x${h}`, 5000).catch(() => {});
        await wait(600);
        const a = await js(AUDIT);
        // A message's body is a frame of its own: measured in it too.
        const inner = await js(FRAMES_AUDIT);
        a.measured += inner.measured;
        for (const b of inner.bad) a.bad.push({ ...b, at: 'message ' + b.at });
        for (const b of a.bad) found.push({ ...b, size: `${w}x${h}` });
        report[app] = { measured: (report[app]?.measured || 0) + a.measured, bad: found };
        if (shots) {
          win.webContents.invalidate();
          await wait(700);
          fs.writeFileSync(path.join(shots, `dark-${app}-${w}x${h}.png`), (await win.webContents.capturePage()).toPNG());
        }
      }
      // The same text at both sizes is one fault.
      const faults = [...new Map(found.map((f) => [`${f.at}|${f.text}`, f])).values()];
      for (const f of faults) console.log(`     [dark-pass] ${app} ${f.size}: ${f.at} "${f.text}" ${f.ink} on ${f.bg} ${f.ratio} (needs ${f.need})`);
      check(`dark-pass: ${app} in the dark theme at 1300x700 and 1680x1050: every word on screen reads on what is under it (4.5:1, large text 3:1)`,
        !faults.length, faults.length ? faults.slice(0, 6).map((f) => `${f.at} "${f.text}" ${f.ink} on ${f.bg} ${f.ratio}`).join('; ') : `${report[app].measured} pieces of text measured`);
    };
    const ready = (win, sel, what) => until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, 15000).catch(() => {});

    const apps = [
      ['home', undefined, '.home-recent, .home-apps, .home'],
      ['word', fx.docx, '.wd-page'],
      ['sheets', fx.xlsx, '.sh-cell'],
      ['slides', fx.pptx, '.sl-thumb'],
      ['pdf', fx.pdf, '.pdf-page canvas[data-drawn]'],
      ['pictures', fx.pics, '.rw-ribbon .rw-tab'],
      ['image', fx.png, '.rw-ribbon .rw-tab'],
      ['video', fx.png, '.vd-clip'],
    ];
    for (const [app, file, sel] of apps.filter(([a]) => wanted(a))) {
      try {
        const win = await open(app, file);
        await ready(win, sel, `the ${app} window`);
        await wait(500);
        await look(app, win);
        const js = (code) => win.webContents.executeJavaScript(code);
        const escape = async (gone) => {
          await js(`document.activeElement?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
          await until(() => js(gone), 'it to close', 4000).catch(() => {});
        };
        // The menus and dialogs every app shares, looked at once each: the app menu from Home, Print from Documents.
        if (app === 'home') {
          await look('app-menu', win, async () => {
            await js(`document.querySelector('.rw-appmark')?.click(), 1`);
            await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the app menu', 4000).catch(() => {});
          });
          await escape(`!document.querySelector('.rw-menu')`);
        }
        if (app === 'word') {
          await look('print', win, async () => {
            const pressed = await js(`(() => { const b = document.querySelector('[data-tip^="Print"]'); b?.click(); return Boolean(b); })()`);
            if (!pressed) win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'P', modifiers: ['control'] });
            await until(() => js(`Boolean(document.querySelector('.rw-dialog'))`), 'the Print dialog', 6000).catch(() => {});
            await wait(600);
          });
          await escape(`!document.querySelector('.rw-dialog')`);
        }
      } catch (err) {
        check(`dark-pass: the ${app} pass ran`, false, err.stack || err.message);
      }
    }

    if (wanted('contacts')) {
      try {
        const win = await open('contacts', fx.vcf);
        await ready(win, '.ct-item', 'the contacts');
        await look('contacts', win, async (js) => {
          await js(`document.querySelector('.ct-item')?.click(), 1`);
          await until(() => js(`Boolean(document.querySelector('.ct-card'))`), "a contact's card", 5000).catch(() => {});
        });
      } catch (err) {
        check('dark-pass: the contacts pass ran', false, err.stack || err.message);
      }
    }

    if (wanted('calendar')) {
      try {
        const win = await open('calendar', fx.ics);
        await ready(win, '.cal-month-grid, .cal-time', 'the calendar');
        await look('calendar', win);
        const js = (code) => win.webContents.executeJavaScript(code);
        await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Week')?.click(), 1`);
        await until(() => js(`document.querySelectorAll('.cal-hours .cal-col').length > 0`), 'the week', 5000).catch(() => {});
        await look('calendar-week', win);
        await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Month')?.click(), 1`);
      } catch (err) {
        check('dark-pass: the calendar pass ran', false, err.stack || err.message);
      }
    }

    if (wanted('mail')) {
      try {
        const win = await open('mail');
        const js = (code) => win.webContents.executeJavaScript(code);
        mailJs = js;
        const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
        await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0 || Boolean(document.querySelector('.rw-empty'))`), 'the Mail window', 8000).catch(() => {});
        const made = await mail('addAccount', { account: { email: 'dark@checks.example', name: 'Dark (checks)', local: true } });
        accountId = made.id;
        const now = Date.now();
        // A newsletter's HTML with colours of its own: dark grey words, a pale box, a link, a quote.
        const html = [
          `From: Lena Weiss <lena.weiss@example.com>`, 'To: dark@checks.example', 'Subject: The autumn letter', `Date: ${new Date(now).toUTCString()}`, 'MIME-Version: 1.0',
          'Content-Type: text/html; charset=utf-8', '',
          '<div style="font-family: Arial; color: #333333"><h2 style="color: #1F3A5F">The autumn letter</h2>'
          + '<p>Three things happened this month. The <a href="https://example.com/range">new range</a> is in the shops.</p>'
          + '<div style="background: #F2F4F7; padding: 10px; border-radius: 6px">Orders placed before Friday ship the same day.</div>'
          + '<blockquote style="color: #666666; border-left: 3px solid #cccccc; padding-left: 8px">We could not be happier with it.</blockquote>'
          + '<p style="color: #999999; font-size: 12px">You are getting this because you asked to.</p></div>',
        ].join('\r\n');
        await mail('deliverTest', { accountId, folder: 'Inbox', raw: html });
        await mail('deliverTest', { accountId, folder: 'Inbox', raw: message({ from: 'Sara Khan <sara.khan@example.com>', subject: 'The spring range', body: 'The samples arrive on Tuesday.', date: new Date(now - 3600e3).toUTCString() }) });
        await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => /Dark \\(checks\\)|dark@checks/.test(n.textContent)); b?.click(); return 1; })()`);
        await until(() => js(`document.querySelectorAll('.ml-row').length >= 2`), 'the two messages', 8000).catch(() => {});
        await look('mail', win, async () => {
          await js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((n) => /autumn letter/.test(n.textContent)); r?.click(); return 1; })()`);
          await until(() => js(`Boolean(document.querySelector('.ml-reader iframe, .ml-body, .ml-reader'))`), 'the message open', 6000).catch(() => {});
          await wait(800);
        });
      } catch (err) {
        check('dark-pass: the mail pass ran', false, err.stack || err.message);
      } finally {
        if (accountId && mailJs) await mailJs(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} })`).catch(() => {});
      }
    }
    console.log(`     [dark-pass] ${JSON.stringify(Object.fromEntries(Object.entries(report).map(([k, v]) => [k, { measured: v.measured, bad: v.bad.length }])))}`);
  } catch (err) {
    check('dark-pass: the checks ran', false, err.stack || err.message);
  } finally {
    if (keep) await keep(`window.rutbaOffice.store.set({ key: 'theme', value: 'light' }).then(() => 1)`).catch(() => {});
  }
}
