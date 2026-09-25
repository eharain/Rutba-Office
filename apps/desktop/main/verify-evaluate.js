// Worksheets: Formulas → Evaluate Formula, and Calculation Options.
//
// C1 adds one to B1, whose IF sums A1:A2 when A1 is the larger. Evaluate
// Formula opens on C1 with B1 underlined; Step In shows B1's own formula
// below it, Evaluate works it out a part at a time — A1, A2, the test, the
// range as {5;3}, the SUM, the product, the IF — and Step Out puts 16 back
// in C1's formula, which ends at 17 with the button reading Restart. Then
// Calculation Options → Manual: a new A1 leaves B1 and C1 where they were
// and the status bar says Calculate; F9 brings them up to date and the word
// goes; Automatic is chosen again. Run alone with RUTBA_VERIFY_ONLY=evaluate.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

const ROWS = [
  [5, '=IF(A1>A2,SUM(A1:A2)*2,0)', '=B1+1'],
  [3],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixture is written
 */
export async function verifyEvaluate(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'evaluate.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Calc', rows: ROWS }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C1"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const menuItem = async (pattern) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)))`), 'the menu', 4000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    };
    const cell = (ref) => `(document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent || '')`;
    const select = async (ref) => {
      await js(`(() => { document.querySelector('.sh-cell[data-ref="${ref}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      return until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === '${ref}'`), ref + ' active', 4000).catch(() => false);
    };
    // What the dialog shows: each level's reference and text, the underlined
    // part and the italic one, and which buttons are live.
    const dialog = () => js(`(() => {
      const levels = [...document.querySelectorAll('.sh-eval-level')].map((l) => ({
        ref: l.querySelector('.sh-eval-ref')?.textContent || '',
        text: l.querySelector('.sh-eval-text')?.textContent || '',
        next: l.querySelector('.sh-eval-next')?.textContent ?? null,
        recent: l.querySelector('.sh-eval-recent')?.textContent ?? null,
      }));
      const button = (cls) => { const b = document.querySelector('.' + cls); return b ? (b.disabled ? 'off' : 'on') : 'none'; };
      return { levels, evaluate: button('sh-eval-evaluate'), restart: button('sh-eval-restart'), stepIn: button('sh-eval-in'), stepOut: button('sh-eval-out'), message: document.querySelector('.sh-eval-message')?.textContent || '' };
    })()`);
    const pressIn = async (cls) => {
      await until(() => js(`Boolean(document.querySelector('.${cls}') && !document.querySelector('.${cls}').disabled)`), cls, 4000).catch(() => {});
      return js(`(() => { const b = document.querySelector('.${cls}'); if (!b || b.disabled) return 'no ${cls}'; b.click(); return 'clicked'; })()`);
    };
    /** The current level's text after the press, waited for. */
    const step = async (cls, expected, level = null) => {
      const before = JSON.stringify(await dialog());
      const pressed = await pressIn(cls);
      await until(async () => JSON.stringify(await dialog()) !== before, 'the dialog to move', 4000).catch(() => {});
      const d = await dialog();
      const at = level === null ? d.levels[d.levels.length - 1] : d.levels[level];
      return { pressed, d, ok: at && at.text === expected };
    };

    // Evaluate Formula on C1.
    const onC1 = await select('C1');
    await tab('Formulas');
    const opened = await clickIn('Evaluate Formula');
    const shown = await until(() => js(`Boolean(document.querySelector('.sh-eval-level'))`), 'the Evaluate Formula dialog', 5000).catch(() => false);
    const first = await dialog();
    check('sheets: Formulas → Evaluate Formula opens on C1 with its reference, its formula and B1 underlined, Step In live',
      onC1 === true && opened === 'clicked' && shown === true && first.levels.length === 1 && first.levels[0].ref === 'Calc!$C$1'
        && first.levels[0].text === 'B1+1' && first.levels[0].next === 'B1' && first.stepIn === 'on' && first.stepOut === 'off' && first.evaluate === 'on',
      JSON.stringify(first));

    // Step In: B1's own formula below, A1 underlined.
    const inB1 = await step('sh-eval-in', 'IF(A1>A2,SUM(A1:A2)*2,0)');
    check('sheets: Step In shows B1\'s own formula below C1\'s, its first reference underlined, Step Out live',
      inB1.ok && inB1.d.levels.length === 2 && inB1.d.levels[1].ref === 'Calc!$B$1' && inB1.d.levels[1].next === 'A1' && inB1.d.stepOut === 'on',
      JSON.stringify(inB1.d));

    // Evaluate through B1, part by part.
    const path1 = [];
    for (const want of ['IF(5>A2,SUM(A1:A2)*2,0)', 'IF(5>3,SUM(A1:A2)*2,0)', 'IF(TRUE,SUM(A1:A2)*2,0)', 'IF(TRUE,SUM({5;3})*2,0)']) {
      const s = await step('sh-eval-evaluate', want);
      path1.push(s.d.levels[1]?.text + (s.ok ? '' : ' (wanted ' + want + ')'));
    }
    const mid = await dialog();
    check('sheets: Evaluate works B1 out a part at a time — A1, A2, the test, then the range shown as {5;3}, the next part underlined and the last result in italics',
      path1.every((t) => !/wanted/.test(t)) && mid.levels[1].next === 'SUM({5;3})' && mid.levels[1].recent === '{5;3}',
      `${path1.join(' → ')}; next ${mid.levels[1]?.next}, recent ${mid.levels[1]?.recent}`);
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'sheets-evaluate.png');

    for (const want of ['IF(TRUE,8*2,0)', 'IF(TRUE,16,0)', '16']) {
      const s = await step('sh-eval-evaluate', want);
      path1.push(s.d.levels[1]?.text + (s.ok ? '' : ' (wanted ' + want + ')'));
    }
    const out = await step('sh-eval-out', '16+1', 0);
    const last = await step('sh-eval-evaluate', '17', 0);
    check('sheets: Step Out puts B1\'s 16 into C1\'s formula, which ends at 17 with the button reading Restart',
      path1.every((t) => !/wanted/.test(t)) && out.ok && out.d.levels.length === 1 && last.ok && last.d.restart === 'on' && /fully evaluated/.test(last.d.message),
      `${path1.slice(4).join(' → ')}; out ${JSON.stringify(out.d.levels)}; last ${JSON.stringify(last.d)}`);
    const restarted = await step('sh-eval-restart', 'B1+1', 0);
    check('sheets: Restart starts C1 over', restarted.ok && restarted.d.levels[0].next === 'B1', JSON.stringify(restarted.d.levels));
    await js(`(() => { [...document.querySelectorAll('.rw-dialog-foot .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(); return 1; })()`);
    await until(() => js(`!document.querySelector('.sh-eval')`), 'the dialog to close', 4000).catch(() => {});

    // Calculation Options → Manual.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith('Calculation Options'))?.click(); return 1; })()`);
    const manualPicked = await menuItem(/^Manual$/);
    const manual = await until(() => model().calc?.mode === 'manual', 'manual mode', 5000).catch(() => false);
    // A new A1: typed as a person types it.
    await select('A1');
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    for (const ch of '10') {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: ch });
      win.webContents.sendInputEvent({ type: 'char', keyCode: ch });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: ch });
      await wait(80);
    }
    await wait(150);
    await press(win.webContents, 'Return');
    const typed = await until(() => js(`${cell('A1')} === '10'`), 'A1 to read 10', 5000).catch(() => false);
    const waiting = await until(() => js(`Boolean(document.querySelector('.sh-calc-pending'))`), 'the status bar to say Calculate', 5000).catch(() => false);
    const stale = await js(`[${cell('B1')}, ${cell('C1')}].join(',')`);
    check('sheets: Calculation Options → Manual — a new A1 leaves B1 and C1 as they were, and the status bar says Calculate',
      manualPicked === 'clicked' && manual === true && typed === true && waiting === true && stale === '16,17',
      `${manualPicked}; manual ${manual}; typed ${typed}; Calculate ${waiting}; B1,C1 ${stale}`);
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'sheets-calc-manual.png');

    // F9: everything up to date, the word gone.
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'F9');
    const caught = await until(() => js(`${cell('B1')} === '26' && ${cell('C1')} === '27' && !document.querySelector('.sh-calc-pending')`), 'F9 to calculate', 5000).catch(() => false);
    check('sheets: F9 (Calculate Now) brings B1 and C1 up to date and the status bar stops saying Calculate',
      caught === true, await js(`[${cell('B1')}, ${cell('C1')}, Boolean(document.querySelector('.sh-calc-pending'))].join(',')`));

    // Calculate Sheet from the ribbon, after another edit.
    await select('A2');
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    for (const ch of '20') {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: ch });
      win.webContents.sendInputEvent({ type: 'char', keyCode: ch });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: ch });
      await wait(80);
    }
    await wait(150);
    await press(win.webContents, 'Return');
    await until(() => js(`${cell('A2')} === '20' && Boolean(document.querySelector('.sh-calc-pending'))`), 'A2 and Calculate', 5000).catch(() => {});
    const sheetPressed = await clickIn('Calculate Sheet');
    const sheetDone = await until(() => js(`${cell('B1')} === '0' && ${cell('C1')} === '1' && !document.querySelector('.sh-calc-pending')`), 'Calculate Sheet', 5000).catch(() => false);
    check('sheets: Formulas → Calculate Sheet works the sheet out — A2 now the larger, the IF answers 0',
      sheetPressed === 'clicked' && sheetDone === true, `${sheetPressed}; ${await js(`[${cell('B1')}, ${cell('C1')}].join(',')`)}`);

    // Automatic again, as the workbook came.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith('Calculation Options'))?.click(); return 1; })()`);
    await menuItem(/^Automatic$/);
    const auto = await until(() => model().calc?.mode === 'auto', 'automatic mode', 5000).catch(() => false);
    check('sheets: Calculation Options → Automatic puts the workbook back', auto === true, JSON.stringify(model().calc));

    const complaints = await errorsIn(win);
    check('sheets: the Evaluate Formula and calculation checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the Evaluate Formula and calculation checks ran', false, err.message);
  }
}
