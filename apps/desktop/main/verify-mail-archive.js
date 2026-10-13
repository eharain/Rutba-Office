// Mail as the archive tool.
//
// Two archives written for the check (a .pst by the suite's own writer and an
// .olm) are opened in place beside an account of its own; the .pst reads in
// the window as a mailbox, read only, with its strip saying so. One search
// over all mail, with the new operators, finds the messages about one matter
// in all three and each result says where it lives; one opens in place. The
// results are exported for legal hold to an mbox and PDFs with the
// attachments beside them, and the files are read back: the mbox's message
// count and Message-IDs, the PDF's pages and a subject's words, and every
// hash in the manifest against the file it names. The archives are unchanged
// at the end. Then the same controls in Urdu and Arabic, right to left.
// Run alone with RUTBA_VERIFY_ONLY=mail-archive.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { Mbox } from '@rutba/mailbox/mbox';
import { parseMessage } from '@rutba/mailbox/mime';
import { writeArchiveFixtures, ACCOUNT_MESSAGES, rfc822 } from '../../../tests/fixtures/mail-archives.js';

const sha = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');
const pdfPages = (bytes) => (Buffer.from(bytes).toString('latin1').match(/\/Type\s*\/Page(?!s)/g) || []).length;

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailArchive(h) {
  const { open, check, until, wait, errorsIn } = h;
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-archive-'));
  const files = writeArchiveFixtures(path.join(root, 'archives'));
  const out = path.join(root, 'export');
  fs.mkdirSync(out);
  const before = Object.fromEntries(Object.entries(files).map(([k, f]) => [k, sha(fs.readFileSync(f))]));
  let win = null;
  let js = null;
  let accountId = null;
  const archiveIds = [];
  let language = null;

  const captureOf = (w) => async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE || !w || w.isDestroyed()) return;
    w.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await w.webContents.capturePage()).toPNG());
  };
  const driver = (w) => {
    const run = (code) => w.webContents.executeJavaScript(code);
    return {
      js: run,
      tab: (name) => run(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`),
      press: (label) =>
        run(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-ribbon button')].find((n) => { const s = (n.dataset.tip || n.title || n.textContent || '').trim(); return s === ${JSON.stringify(label)} || s.startsWith(${JSON.stringify(`${label}:`)}) || n.textContent.trim() === ${JSON.stringify(label)}; }); if (!b) return 'no button ' + ${JSON.stringify(label)}; if (b.disabled) return 'disabled'; b.click(); return 'pressed'; })()`),
      type: (selector, value) =>
        run(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return 'no field'; const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true })); return 'typed'; })()`),
      clickText: (selector, text) =>
        run(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(selector)})].find((n) => n.textContent.trim().startsWith(${JSON.stringify(text)})); if (!b) return 'no ' + ${JSON.stringify(text)}; b.click(); return 'clicked'; })()`),
    };
  };

  /** Wait for `fn` to give something, and give it back (until() answers only true). */
  const waitFor = async (fn, what, ms) => {
    let got = null;
    await until(async () => {
      got = await fn();
      return Boolean(got);
    }, what, ms).catch(() => {});
    return got || null;
  };

  try {
    win = await open('mail');
    const d = driver(win);
    js = d.js;
    const capture = captureOf(win);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    language = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the sidebar', 8000);

    // The account fixture: an account of the check's own, removed at the end.
    const made = await mail('addAccount', { account: { email: 'legal@checks.example', name: 'Legal (checks)', local: true } });
    accountId = made.id;
    for (const m of ACCOUNT_MESSAGES) await mail('deliverTest', { accountId, folder: 'Inbox', raw: rfc822(m) });

    // Two archives opened in place: the sidebar lists them, as it is told.
    for (const f of [files.pst, files.olm]) archiveIds.push((await mail('openArchive', { path: f })).id);
    const listed = await waitFor(
      async () => {
        const names = await js(`[...document.querySelectorAll('.ml-archives button')].map((n) => n.textContent.trim())`);
        return names.some((n) => n.startsWith('Northwind 2019')) && names.some((n) => n.startsWith('Contoso archive')) ? names : false;
      },
      'both archives in the sidebar',
      8000
    );
    check('mail archive: two archives opened in place are listed under Archives in the sidebar', Array.isArray(listed) && listed.length >= 3, JSON.stringify(listed));

    // The .pst reads as a mailbox: its folders, its Inbox, a message; read only.
    await d.clickText('.ml-archives button', 'Northwind 2019');
    const shown = await waitFor(
      async () => {
        const s = await js(`(() => ({
          banner: document.querySelector('.ml-archive-banner')?.textContent || '',
          folders: [...document.querySelectorAll('.rw-list button')].map((n) => n.textContent.trim()).filter(Boolean),
          rows: document.querySelectorAll('.ml-row').length,
        }))()`);
        return /Read only/.test(s.banner) && s.folders.some((f) => /^Projects/.test(f)) && s.rows > 0 ? s : false;
      },
      'the archive drawn as a mailbox',
      10000
    );
    check('mail archive: the .pst opens in the window with its folders, its list and a strip saying it is read only', Boolean(shown), JSON.stringify(shown));
    await d.clickText('.ml-row .ml-subject', 'Re: Halvorsen contract');
    const reading = await waitFor(async () => {
      const h2 = await js(`document.querySelector('.ml-head h2')?.textContent || ''`);
      return /Halvorsen contract/.test(h2) ? h2 : false;
    }, 'a message from the archive in the reading pane', 6000);
    const stillUnread = await js(`Boolean([...document.querySelectorAll('.ml-row')].find((r) => r.querySelector('.ml-subject')?.textContent.includes('Re: Halvorsen contract'))?.classList.contains('unread'))`);
    const deleteDisabled = await js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Delete')?.disabled)`);
    check('mail archive: a message opens in place, is not marked read, and Delete is not offered', Boolean(reading) && stillUnread && deleteDisabled, JSON.stringify({ reading, stillUnread, deleteDisabled }));
    await capture('mail-archive-1-open-in-place.png');

    // Indexed, both of them, before the search across all.
    const ready = await waitFor(async () => {
      const list = await mail('archives');
      return list.length >= 2 && list.every((a) => a.index.state === 'ready') ? list : false;
    }, 'the archives indexed', 15000);
    check('mail archive: both archives are indexed for search', Boolean(ready), JSON.stringify((await mail('archives')).map((a) => [a.name, a.index.state, a.index.done, a.index.total])));

    // One search over all mail, with the new operators.
    const searchFor = async (query) => {
      await d.type('.rw-search input', query);
      await until(() => js(`Boolean(document.querySelector('.ml-scope'))`), 'the search scope bar', 4000).catch(() => {});
      await d.clickText('.ml-scope-seg button', 'All mail');
      return waitFor(async () => {
        const s = await js(`(() => { const f = document.querySelector('.ml-scope-found'); return f ? { total: Number(f.dataset.total), sources: Number(f.dataset.sources), text: f.textContent, where: [...document.querySelectorAll('.ml-where span')].map((n) => n.textContent), subjects: [...document.querySelectorAll('.ml-row .ml-subject')].map((n) => n.textContent.trim()), query: f.dataset.query } : null; })()`);
        return s && s.query === query ? s : false;
      }, `the results of ${query}`, 8000);
    };
    await d.tab('Search');
    const first = await searchFor('halvorsen after:2019-01-01');
    const sources = new Set((first?.where || []).map((w) => w.split(' · ')[0]));
    check(
      'mail archive: one search across all mail (after:) finds five messages in three places, each saying where it lives',
      first && first.total === 5 && first.sources === 3 && sources.has('Northwind 2019') && sources.has('Contoso archive') && sources.has('Legal (checks)') && first.where.some((w) => w === 'Northwind 2019 · Inbox/Projects'),
      JSON.stringify(first)
    );
    await capture('mail-archive-2-search-all.png');
    const second = await searchFor('"indemnity clause" -from:bo');
    check('mail archive: a quoted phrase in the bodies, less a sender, across all mail', second && second.total === 2 && second.sources === 2, JSON.stringify(second));
    const third = await searchFor('filename:pdf has:attachment before:2020-12-31');
    check('mail archive: filename:, has:attachment and before: together', third && third.total === 3, JSON.stringify(third));

    // A result opens in place, from the archive it lives in.
    await searchFor('halvorsen after:2019-01-01');
    await d.clickText('.ml-row .ml-subject', 'Halvorsen site visit');
    const visit = await waitFor(async () => {
      const h2 = await js(`document.querySelector('.ml-head h2')?.textContent || ''`);
      return h2.includes('Halvorsen site visit') ? h2 : false;
    }, 'the .olm result in the reading pane', 6000);
    check('mail archive: a result from the .olm opens in the reading pane', Boolean(visit), visit);

    // The search language's help lists the operators.
    await d.press('Search syntax');
    const help = await waitFor(() => js(`document.querySelectorAll('.ml-help-row').length || false`), 'the search syntax', 4000);
    await capture('mail-archive-3-search-syntax.png');
    check('mail archive: Search syntax lists the operators', help >= 14, String(help));
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await until(() => js(`!document.querySelector('.ml-help-row')`), 'the help closed', 3000).catch(() => {});

    // Export the results for legal hold: an mbox, PDFs, the attachments beside them.
    await searchFor('halvorsen after:2019-01-01');
    const pressed = await d.press('Export results');
    await until(() => js(`Boolean(document.querySelector('.ml-hold-section'))`), 'the export dialog', 5000).catch(() => {});
    await d.type('.ml-hold select', 'both');
    await js(`(() => { const box = [...document.querySelectorAll('.ml-hold-check')].find((l) => l.textContent.includes('attachments')); const i = box?.querySelector('input'); if (i && !i.checked) i.click(); return Boolean(i?.checked); })()`);
    await d.type('.ml-hold-dir input', out);
    await wait(200);
    await capture('mail-archive-4-export-dialog.png');
    const go = await js(`(() => { const b = [...document.querySelectorAll('.rw-dialog-foot button')].find((n) => /^Export \\d+ message/.test(n.textContent.trim())); if (!b) return 'no button'; if (b.disabled) return 'disabled: ' + b.textContent; b.click(); return b.textContent.trim(); })()`);
    const sentence = await waitFor(async () => {
      const s = await js(`document.querySelector('.ml-hold-sentence')?.textContent || ''`);
      return s || false;
    }, 'the export to finish', 30000);
    await capture('mail-archive-5-export-done.png');
    check('mail archive: Export results exports the five, and ends with a sentence of what was written', pressed === 'pressed' && /Export 5 messages/.test(go) && /^Exported 5 messages from 3 mailboxes and archives/.test(sentence) && /an mbox/.test(sentence) && /Nothing in the mailboxes or archives was changed/.test(sentence), JSON.stringify({ pressed, go, sentence }));
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);

    // The files, read back.
    const made2 = fs.readdirSync(out).map((n) => path.join(out, n)).find((p) => fs.statSync(p).isDirectory());
    const manifest = made2 ? JSON.parse(fs.readFileSync(path.join(made2, 'manifest.json'), 'utf8')) : null;
    const mbox = made2 ? new Mbox(fs.readFileSync(path.join(made2, 'messages.mbox'))) : null;
    const ids = [];
    const hashOk = [];
    for (let i = 0; mbox && i < mbox.length; i++) {
      const raw = Buffer.from(mbox.raw(i)).subarray(0, -1);
      ids.push(parseMessage(raw).messageId);
      hashOk.push(sha(raw) === manifest.messages[i].sha256);
    }
    check(
      'mail archive: the mbox reads back with five messages, their Message-IDs, and the hash the manifest gives each',
      mbox?.length === 5 && hashOk.every(Boolean) && ['halvorsen-1@northwind.example', 'site-visit@contoso.example', 'halvorsen-next@northwind.example'].every((id) => ids.includes(id)),
      JSON.stringify({ count: mbox?.length, ids, hashOk })
    );
    const combined = made2 ? fs.readFileSync(path.join(made2, 'messages.pdf')) : Buffer.alloc(0);
    const eachPdf = made2 ? fs.readdirSync(path.join(made2, 'pdf')) : [];
    check(
      'mail archive: the PDF of them all has a summary page and a page or more each, and shows a subject',
      pdfPages(combined) >= 6 && combined.toString('latin1').includes('(Halvorsen site visit) Tj') && eachPdf.length === 5,
      JSON.stringify({ pages: pdfPages(combined), each: eachPdf })
    );
    const mismatched = [];
    for (const f of manifest?.files || []) if (sha(fs.readFileSync(path.join(made2, f.name))) !== f.sha256) mismatched.push(f.name);
    for (const m of manifest?.messages || []) {
      if (m.pdf && sha(fs.readFileSync(path.join(made2, m.pdf.file))) !== m.pdf.sha256) mismatched.push(m.pdf.file);
      for (const a of m.attachments) if (a.file && sha(fs.readFileSync(path.join(made2, a.file))) !== a.sha256) mismatched.push(a.file);
    }
    const attachmentsWritten = (manifest?.messages || []).flatMap((m) => m.attachments).filter((a) => a.file).length;
    check('mail archive: every hash in the manifest matches the file it names', manifest && mismatched.length === 0 && attachmentsWritten >= 1 && fs.existsSync(path.join(made2, 'manifest.csv')) && fs.existsSync(path.join(made2, 'summary.pdf')), JSON.stringify({ mismatched, attachmentsWritten, files: manifest?.files?.map((f) => f.name) }));

    const after = Object.fromEntries(Object.entries(files).map(([k, f]) => [k, sha(fs.readFileSync(f))]));
    check('mail archive: the archives are byte for byte as they were', JSON.stringify(after) === JSON.stringify(before), JSON.stringify(Object.keys(files)));
    const complaints = await errorsIn(win);
    check('mail archive: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // Right to left: the Search tab, the scope bar, the strip and the export dialog in Urdu and Arabic.
    for (const tag of ['ur', 'ar']) {
      await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(tag)} }).then(() => 1)`);
      const w = await open('mail');
      const dw = driver(w);
      await until(() => dw.js(`document.querySelectorAll('.ml-archives button').length >= 2`), `the archives in ${tag}`, 8000).catch(() => {});
      await dw.js(`document.querySelectorAll('.ml-archives button')[0]?.click(), 1`);
      await until(() => dw.js(`Boolean(document.querySelector('.ml-archive-banner'))`), `the strip in ${tag}`, 6000).catch(() => {});
      const tabs = await dw.js(`[...document.querySelectorAll('.rw-tab')].map((n) => n.textContent.trim())`);
      await dw.js(`document.querySelectorAll('.rw-tab')[1]?.click(), 1`);
      await dw.type('.rw-search input', 'halvorsen after:2019-01-01');
      await until(() => dw.js(`document.querySelectorAll('.ml-scope-seg button').length === 2`), `the scope bar in ${tag}`, 4000).catch(() => {});
      await dw.js(`document.querySelectorAll('.ml-scope-seg button')[1]?.click(), 1`);
      await until(() => dw.js(`Number(document.querySelector('.ml-scope-found')?.dataset.total) === 5`), `the results in ${tag}`, 8000).catch(() => {});
      const frame = await dw.js(`(() => {
        const r = (n) => n?.getBoundingClientRect();
        const seg = [...document.querySelectorAll('.ml-scope-seg button')];
        const ribbon = [...document.querySelectorAll('.rw-ribbon .rw-group')];
        const banner = document.querySelector('.ml-archive-banner');
        return {
          dir: document.documentElement.dir,
          segRtl: seg.length === 2 ? r(seg[0]).left > r(seg[1]).left : false,
          ribbonRtl: ribbon.length > 1 ? r(ribbon[0]).left > r(ribbon[1]).left : false,
          bannerRtl: banner ? r(banner.querySelector('svg')).left > r(banner.querySelector('.ml-archive-name')).left : false,
          words: [...document.querySelectorAll('.rw-ribbon .rw-group .rw-btn, .ml-scope-seg button, .ml-archive-banner button, .ml-archive-note, .ml-scope-found')].map((n) => n.textContent.trim()).filter(Boolean),
        };
      })()`);
      await captureOf(w)(`mail-archive-6-${tag}-search.png`);
      // The export dialog too.
      const exportLabel = await dw.js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-group')][3]?.querySelector('.rw-btn'); if (!b || b.disabled) return 'none'; b.click(); return b.textContent.trim(); })()`);
      await until(() => dw.js(`Boolean(document.querySelector('.ml-hold-section'))`), `the export dialog in ${tag}`, 4000).catch(() => {});
      const dialogWords = await dw.js(`[...document.querySelectorAll('.ml-hold-section, .ml-hold-choice span, .ml-hold-check span, .rw-dialog-head')].map((n) => n.textContent.trim())`);
      await captureOf(w)(`mail-archive-7-${tag}-export.png`);
      const english = [...frame.words, ...dialogWords].filter((s) => /^(Search|Scope|Refine|This folder|All mail|From|Sent to|Subject|Has attachments|Attachment name|Unread|Date|Size|Open archive|Close archive|Export results|Export folders|Search syntax|Export|What to export|Write|Export for legal hold|Close archive)$/.test(s) || /^\d+ found/.test(s));
      check(
        `mail archive: in ${tag === 'ur' ? 'Urdu' : 'Arabic'} the Search tab, the scope bar, the archive's strip and the export dialog read right to left in that language`,
        frame.dir === 'rtl' && frame.segRtl && frame.ribbonRtl && frame.bannerRtl && english.length === 0 && dialogWords.length >= 4 && tabs.length >= 5 && exportLabel !== 'none',
        JSON.stringify({ ...frame, words: frame.words.slice(0, 12), dialogWords, english, exportLabel })
      );
      w.destroy();
    }
  } catch (err) {
    check('mail archive: the checks ran', false, err.stack || err.message);
  } finally {
    // Leave everything as it was: the language, the archives closed, the account gone.
    try {
      if (js) {
        await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(language ?? null)} }).then(() => 1)`).catch(() => {});
        for (const id of archiveIds) await js(`window.rutbaOffice.mail.closeArchive({ id: ${JSON.stringify(id)} })`).catch(() => {});
        if (accountId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)}, keepMail: false })`).catch(() => {});
      }
    } finally {
      try {
        fs.rmSync(root, { recursive: true, force: true });
      } catch {
        /* a window may still hold a file */
      }
    }
  }
}
