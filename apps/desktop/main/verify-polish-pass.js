// A visual pass over every app: each window at 1300x700 and 1680x1050, in
// the light theme and the dark, written to RUTBA_VERIFY_CAPTURE as
// polish-<app>-<size>-<theme>.png to be looked at. Measured at each: the
// page never scrolls sideways and the ribbon keeps its rules (nothing cut
// off, nothing past the window, the folded groups last). And two faults
// found by eye stay put right: Mail's empty reading pane names its keys
// with a space after each comma, and Calendar's week and day views keep
// their day columns one width, the headings over the hours, with a busy
// day and a long all-day title. Its files, and its Mail account, are its
// own and go when it is done. Run alone with RUTBA_VERIFY_ONLY=polish.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import pdfWriter from '@rutba/pdf';
import { gradientPng } from './sample-picture.js';
import { MEASURE } from './verify-ribbon.js';

const { PdfDocument } = pdfWriter;
const SIZES = [[1300, 700], [1680, 1050]];
/** RUTBA_POLISH_APPS=mail,calendar looks at those apps alone, for working on them. */
const wanted = (app) => { const only = (process.env.RUTBA_POLISH_APPS || '').split(',').map((x) => x.trim()).filter(Boolean); return !only.length || only.includes(app); };

const pad = (n) => String(n).padStart(2, '0');
const icsDate = (d) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
const icsTime = (d, h, m = 0) => `${icsDate(d)}T${pad(h)}${pad(m)}00`;

function fixtures(dir) {
  const own = path.join(dir, 'polish-fixtures');
  const pics = path.join(own, 'pictures');
  fs.mkdirSync(pics, { recursive: true });
  const at = (n) => path.join(own, n);
  fs.writeFileSync(at('polish.docx'), buildDocx({ styles: true, paragraphs: [
    { text: 'Quarterly review', style: 'Title' },
    { text: 'Where the year stands', style: 'Heading1' },
    { text: 'Sales rose in every region but one. The north held steady while the coast grew by a fifth, and the new shop in the city paid for itself before the summer.' },
    { text: 'What comes next', style: 'Heading2' },
    { text: 'Two new lines in the spring, and the warehouse moved closer to the port.' },
  ] }));
  fs.writeFileSync(at('polish.xlsx'), buildXlsx({ sheets: [{ name: 'Sales', rows: [
    ['Region', 'Q1', 'Q2', 'Q3', 'Q4'], ['North', 120, 132, 128, 140], ['Coast', 98, 110, 126, 141], ['City', 40, 76, 92, 118], ['Hills', 64, 61, 70, 72],
  ] }] }));
  fs.writeFileSync(at('polish.pptx'), buildPptx({ title: 'Polish', slides: [
    { layout: 'title', title: 'Quarterly review', body: 'The year so far' }, { layout: 'obj', title: 'Where we stand', body: ['Sales up in three regions', 'A new shop in the city', 'The warehouse moves in spring'] },
  ] }));
  const pdf = new PdfDocument({ title: 'polish', created: '2026-10-14T09:00:00Z' });
  for (const title of ['Quarterly review', 'Where we stand']) pdf.addPage(null).text(title, 72, 110, { font: 'Helvetica-Bold', size: 30 });
  fs.writeFileSync(at('polish.pdf'), pdf.toBuffer());
  const landscapes = [[[20, 110, 140], [240, 190, 90]], [[40, 60, 120], [200, 120, 160]], [[30, 120, 70], [230, 220, 140]]];
  landscapes.forEach(([a, b], i) => fs.writeFileSync(path.join(pics, `view-${i + 1}.png`), gradientPng(960, 640, a, b)));
  fs.writeFileSync(at('polish.png'), gradientPng(1280, 800, [20, 110, 140], [240, 190, 90]));
  // Today busy: five meetings, two of them at once, and an all-day event with a long title.
  const today = new Date();
  const ev = (uid, summary, start, end, allDay = false) => [
    'BEGIN:VEVENT', `UID:${uid}@polish.checks`, `DTSTAMP:${icsTime(today, 8)}`,
    allDay ? `DTSTART;VALUE=DATE:${icsDate(start)}` : `DTSTART:${start}`,
    allDay ? `DTEND;VALUE=DATE:${icsDate(end)}` : `DTEND:${end}`,
    `SUMMARY:${summary}`, 'END:VEVENT',
  ].join('\r\n');
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  fs.writeFileSync(at('polish.ics'), ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//polish checks//EN',
    ev('a', 'Planning with the regional managers about the spring range and the warehouse move', today, tomorrow, true),
    ev('b', 'Stand-up', icsTime(today, 9), icsTime(today, 9, 30)),
    ev('c', 'Budget review for the coming year', icsTime(today, 10), icsTime(today, 11, 30)),
    ev('d', 'Call with the printers', icsTime(today, 10, 30), icsTime(today, 11)),
    ev('e', 'Lunch', icsTime(today, 12, 30), icsTime(today, 13, 30)),
    ev('f', 'Interview: shop manager', icsTime(today, 15), icsTime(today, 16)),
    'END:VCALENDAR'].join('\r\n'));
  fs.writeFileSync(at('polish.vcf'), [
    ['Sara Khan', 'Khan;Sara', 'sara.khan@example.com', '+44 20 7946 0001', 'Northwind'],
    ['Omar Malik', 'Malik;Omar', 'omar@example.com', '+44 20 7946 0002', 'Contoso'],
    ['Lena Weiss', 'Weiss;Lena', 'lena.weiss@example.com', '+49 30 901820', 'Fabrikam'],
  ].map(([fn, n, mail, tel, org]) => `BEGIN:VCARD\r\nVERSION:3.0\r\nFN:${fn}\r\nN:${n};;;\r\nEMAIL:${mail}\r\nTEL:${tel}\r\nORG:${org}\r\nEND:VCARD`).join('\r\n'));
  return {
    docx: at('polish.docx'), xlsx: at('polish.xlsx'), pptx: at('polish.pptx'), pdf: at('polish.pdf'), png: at('polish.png'),
    pics, ics: at('polish.ics'), vcf: at('polish.vcf'),
  };
}

/** A message as it arrives, for the check's own account. */
const message = ({ from, subject, body, date }) => [
  `From: ${from}`, 'To: polish@checks.example', `Subject: ${subject}`, `Date: ${date}`, 'MIME-Version: 1.0',
  'Content-Type: text/plain; charset=utf-8', '', body,
].join('\r\n');

export async function verifyPolishPass({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const fx = fixtures(dir);
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };
  const size = async (win, js, w, h) => {
    win.setContentSize(w, h);
    await until(() => js('[innerWidth, innerHeight]').then(([iw, ih]) => Math.abs(iw - w) <= 2 && Math.abs(ih - h) <= 2), `the window at ${w}x${h}`, 5000).catch(() => {});
    await wait(450);
  };
  const theme = (js, name) => js(`document.documentElement.dataset.theme = ${JSON.stringify(name)}, 1`);
  const ok = (m) => m.none || (m.past === 0 && m.over <= 1 && !m.clipped.length && /^[-CS]*F*$/.test(m.stages));
  const sideways = (js) => js(`Math.max(document.documentElement.scrollWidth, document.body.scrollWidth) - innerWidth`);
  // A panel whose body scrolls sideways is cramped: named by its widest child.
  const panelsWide = (js) => js(`[...document.querySelectorAll('.rw-panel-body')].filter((b) => b.getBoundingClientRect().width > 0 && b.scrollWidth > b.clientWidth + 1).map((b) => {
    const box = b.getBoundingClientRect();
    const wide = [...b.querySelectorAll('*')].filter((n) => n.getBoundingClientRect().right > box.right + 1).sort((x, y) => y.getBoundingClientRect().right - x.getBoundingClientRect().right)[0];
    return (b.closest('.rw-panel')?.querySelector('.rw-panel-head')?.textContent.trim() || 'panel') + ': ' + (b.scrollWidth - b.clientWidth) + ' px, ' + (wide ? (wide.className || wide.tagName) : '?');
  })`);

  /** Every size and theme of one window: captures, and the page and ribbon measured. */
  const pass = async (app, win, js, extra = null) => {
    const bad = [];
    for (const [w, h] of SIZES) {
      await size(win, js, w, h);
      for (const name of ['light', 'dark']) {
        await theme(js, name);
        await wait(200);
        const m = await js(MEASURE);
        const over = await sideways(js);
        if (!ok(m)) bad.push(`${w}x${h} ${name}: ribbon past ${m.past}, over ${m.over}, stages ${m.stages}, clipped [${m.clipped.slice(0, 3).join('|')}]`);
        if (over > 1) bad.push(`${w}x${h} ${name}: the page scrolls ${over} px sideways`);
        const cramped = await panelsWide(js);
        if (cramped.length) bad.push(`${w}x${h} ${name}: a panel scrolls sideways (${cramped.join('; ')})`);
        await snap(win, `polish-${app}-${w}x${h}-${name}.png`);
        if (extra) await extra(w, h, name);
      }
    }
    await theme(js, 'light');
    check(`polish: ${app} at 1300x700 and 1680x1050, light and dark, never scrolls sideways, nor do its panels, and its ribbon keeps its rules`, !bad.length, bad.slice(0, 4).join('; ') || 'four views');
  };

  const apps = [
    ['word', fx.docx, `Boolean(document.querySelector('.wd-page'))`],
    ['sheets', fx.xlsx, `Boolean(document.querySelector('.sh-cell'))`],
    ['slides', fx.pptx, `document.querySelectorAll('.sl-thumb').length >= 2`],
    ['pdf', fx.pdf, `Boolean(document.querySelector('.rw-ribbon .rw-tab'))`],
    ['contacts', fx.vcf, `Boolean(document.querySelector('.rw-ribbon .rw-tab'))`],
    ['pictures', fx.pics, `Boolean(document.querySelector('.rw-ribbon .rw-tab, .pv-grid, .pv-strip'))`],
    ['image', fx.png, `Boolean(document.querySelector('.rw-ribbon .rw-tab'))`],
    ['video', fx.png, `document.querySelectorAll('.vd-clip').length > 0`],
    ['home', undefined, `Boolean(document.querySelector('.home-recent, .home-apps, .home'))`],
  ];
  for (const [app, file, ready] of apps.filter(([app]) => wanted(app))) {
    try {
      const win = await open(app, file);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(ready), `the ${app} window`, 15000).catch(() => {});
      await wait(400);
      await pass(app, win, js);
    } catch (err) {
      check(`polish: the ${app} pass ran`, false, err.stack || err.message);
    }
  }

  // Mail, on an account of the pass's own with a few messages and none chosen.
  let accountId = null;
  if (wanted('mail')) {
  let mailJs = null;
  try {
    const win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    mailJs = js;
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0 || Boolean(document.querySelector('.rw-empty'))`), 'the Mail window', 8000).catch(() => {});
    const made = await mail('addAccount', { account: { email: 'polish@checks.example', name: 'Polish (checks)', local: true } });
    accountId = made.id;
    const now = Date.now();
    const notes = [
      { from: 'Sara Khan <sara.khan@example.com>', subject: 'The spring range', body: 'The samples arrive on Tuesday. Can we look at them together?' },
      { from: 'Omar Malik <omar@example.com>', subject: 'Warehouse move: the plan', body: 'Attached is the plan for the move, week by week.' },
      { from: 'Lena Weiss <lena.weiss@example.com>', subject: 'Invoice 2041', body: 'Thank you for the order. The invoice is below.' },
    ];
    for (const [i, n] of notes.entries()) await mail('deliverTest', { accountId, folder: 'Inbox', raw: message({ ...n, date: new Date(now - i * 3600e3).toUTCString() }) });
    await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => /Polish \\(checks\\)|polish@checks/.test(n.textContent)); b?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.ml-row, .ml-item, [data-message]').length >= 3`), 'the three messages', 8000).catch(() => {});
    const hint = await until(() => js(`(() => { const e = [...document.querySelectorAll('.rw-empty')].find((n) => /No message selected/.test(n.textContent)); return e ? e.textContent.replace(/\\s+/g, ' ').trim() : ''; })()`), 'the empty reading pane', 6000).catch(() => '');
    const text = typeof hint === 'string' ? hint : await js(`(() => { const e = [...document.querySelectorAll('.rw-empty')].find((n) => /No message selected/.test(n.textContent)); return e ? e.textContent.replace(/\\s+/g, ' ').trim() : ''; })()`);
    check('polish: Mail\'s empty reading pane names its keys with a space after each comma (j and k move, r replies, s stars)', /j and k move, r replies, s stars\./.test(text), JSON.stringify(text));
    await pass('mail', win, js);
  } catch (err) {
    check('polish: the Mail pass ran', false, err.stack || err.message);
  } finally {
    if (accountId && mailJs) await mailJs(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} })`).catch(() => {});
  }
  }

  // Calendar: month, week and day, with today busy.
  if (wanted('calendar')) try {
    const win = await open('calendar', fx.ics);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.cal-month-grid, .cal-time'))`), 'the calendar', 10000).catch(() => {});
    const view = async (label) => {
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); b?.click(); return b ? 'clicked' : 'none'; })()`);
      await wait(400);
    };
    const columns = `(() => {
      const heads = [...document.querySelectorAll('.cal-time-head .cal-col-head')].map((n) => n.getBoundingClientRect());
      const cols = [...document.querySelectorAll('.cal-hours .cal-col')].map((n) => n.getBoundingClientRect());
      const allday = [...document.querySelectorAll('.cal-allday .cal-allday-cell')].map((n) => n.getBoundingClientRect());
      const widths = cols.map((r) => Math.round(r.width));
      const aligned = heads.length === cols.length && heads.every((h, i) => Math.abs(h.left - cols[i].left) <= 1 && Math.abs(h.right - cols[i].right) <= 1)
        && allday.length === cols.length && allday.every((a, i) => Math.abs(a.left - cols[i].left) <= 1 && Math.abs(a.right - cols[i].right) <= 1);
      const spill = [...document.querySelectorAll('.cal-allday-cell .cal-chip, .cal-col .cal-block')].filter((c) => { const p = c.closest('.cal-allday-cell, .cal-col').getBoundingClientRect(); const r = c.getBoundingClientRect(); return r.right > p.right + 1 || r.left < p.left - 1; }).length;
      return { count: cols.length, widths, equal: widths.length > 0 && Math.max(...widths) - Math.min(...widths) <= 1, aligned, spill };
    })()`;
    const bad = [];
    for (const label of ['Week', 'Day']) {
      await view(label);
      await until(() => js(`document.querySelectorAll('.cal-hours .cal-col').length > 0`), `the ${label} view`, 4000).catch(() => {});
      for (const [w, h] of SIZES) {
        await size(win, js, w, h);
        for (const name of ['light', 'dark']) {
          await theme(js, name);
          await wait(200);
          const c = await js(columns);
          if (!c.equal || !c.aligned || c.spill) bad.push(`${label} ${w}x${h} ${name}: ${JSON.stringify(c)}`);
          await snap(win, `polish-calendar-${label.toLowerCase()}-${w}x${h}-${name}.png`);
        }
      }
      await theme(js, 'light');
    }
    check('polish: Calendar\'s week and day views keep their day columns one width, the headings and all-day row over the hours, nothing spilling out of its day, with today busy',
      !bad.length, bad.slice(0, 3).join('; ') || 'week and day, both sizes, both themes');
    await view('Month');
    await pass('calendar', win, js);
  } catch (err) {
    check('polish: the Calendar pass ran', false, err.stack || err.message);
  }
}
