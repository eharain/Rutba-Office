// Where the time goes when you type.
//
// The document service is the thing between a keystroke and the screen, so this
// times it directly — no window, no IPC — and reports the cost of each part
// separately. A keystroke has a budget of about 16 ms if typing is to feel
// immediate; anything here that eats it is the thing to fix.
//
//   node tools/bench-edit.js [paragraphs] [rows]

import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';

const paragraphs = Number(process.argv[2] || 400);
const rows = Number(process.argv[3] || 2000);

const ms = (fn, times = 1) => {
  const started = process.hrtime.bigint();
  let out;
  for (let i = 0; i < times; i++) out = fn(i);
  const total = Number(process.hrtime.bigint() - started) / 1e6;
  return { each: total / times, total, out };
};

const bar = (value, budget = 16) => {
  const n = Math.min(40, Math.round((value / budget) * 20));
  return `${'█'.repeat(Math.max(0, n))}${value > budget ? ' ← over budget' : ''}`;
};

console.log(`\nDOCUMENT — ${paragraphs} paragraphs\n${'─'.repeat(64)}`);
{
  const bytes = buildDocx({
    styles: true,
    paragraphs: Array.from({ length: paragraphs }, (_, i) =>
      i % 12 === 0
        ? { text: `Section ${i / 12 + 1}`, style: 'Heading1', bold: true, size: 18 }
        : { text: `Paragraph ${i}. The quick brown fox jumps over the lazy dog, and keeps jumping until the line has to wrap somewhere sensible.` }
    ),
  });
  console.log(`file            ${(bytes.length / 1024).toFixed(0)} KB`);

  const open = ms(() => openDocx(bytes));
  const view = open.out;
  console.log(`open            ${open.each.toFixed(1)} ms`);

  const render = ms(() => view.render(), 20);
  const frame = render.out;
  console.log(`render()        ${render.each.toFixed(2)} ms  ${bar(render.each)}`);
  console.log(`  blocks        ${frame.blocks.length}`);
  console.log(`  pages         ${Array.isArray(frame.pages) ? frame.pages.length : typeof frame.pages}`);

  const serialise = ms(() => JSON.stringify(frame), 10);
  console.log(`  as JSON       ${serialise.each.toFixed(2)} ms, ${(serialise.out.length / 1024).toFixed(0)} KB per keystroke  ${bar(serialise.each)}`);

  view.setSelection({ block: 1, offset: 5 }, { block: 1, offset: 5 });
  const type = ms(() => view.insertText('x'), 50);
  console.log(`insertText      ${type.each.toFixed(2)} ms  ${bar(type.each)}`);

  const whole = ms(() => {
    view.insertText('y');
    return JSON.stringify(view.render());
  }, 20);
  console.log(`ROUND TRIP      ${whole.each.toFixed(2)} ms per keystroke  ${bar(whole.each)}`);
}

console.log(`\nTHE ROUND TRIP AS THE WINDOW SEES IT — ${paragraphs} paragraphs\n${'─'.repeat(64)}`);
{
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const service = createDocumentService({ holdBlob: () => ({ url: '' }) });
  const opened = service.new({ kind: 'word' });

  // Fill it, so the measurement is of a document somebody would actually have.
  for (let i = 0; i < paragraphs; i++) {
    service.apply({ id: opened.id, ops: [{ op: 'insertText', text: `Paragraph ${i}. The quick brown fox jumps over the lazy dog.` }, { op: 'splitParagraph' }] });
  }

  const full = ms(() => JSON.stringify(service.apply({ id: opened.id, ops: [{ op: 'insertText', text: 'x' }], delta: false })), 10);
  console.log(`whole model     ${full.each.toFixed(2)} ms, ${(full.out.length / 1024).toFixed(0)} KB per keystroke  ${bar(full.each)}`);

  const delta = ms(() => JSON.stringify(service.apply({ id: opened.id, ops: [{ op: 'insertText', text: 'y' }] })), 10);
  console.log(`patch           ${delta.each.toFixed(2)} ms, ${(delta.out.length / 1024).toFixed(1)} KB per keystroke  ${bar(delta.each)}`);
  console.log(`                ${(full.out.length / Math.max(1, delta.out.length)).toFixed(0)}× less crossing the process boundary`);
}

console.log(`\nWORKSHEET — ${rows} rows\n${'─'.repeat(64)}`);
{
  const bytes = buildXlsx({
    sheets: [
      {
        name: 'Data',
        rows: Array.from({ length: rows }, (_, r) => [`Row ${r}`, r, r * 2, `=B${r + 1}+C${r + 1}`, 'note']),
      },
    ],
  });
  console.log(`file            ${(bytes.length / 1024).toFixed(0)} KB`);

  const open = ms(() => SheetView.open(bytes, { viewportWidth: 1200, viewportHeight: 700 }));
  const view = open.out;
  console.log(`open            ${open.each.toFixed(1)} ms`);

  const render = ms(() => view.render(), 20);
  const frame = render.out;
  console.log(`render()        ${render.each.toFixed(2)} ms  ${bar(render.each)}`);
  console.log(`  cells         ${frame.cells.length}`);

  const serialise = ms(() => JSON.stringify(frame), 10);
  console.log(`  as JSON       ${serialise.each.toFixed(2)} ms, ${(serialise.out.length / 1024).toFixed(0)} KB  ${bar(serialise.each)}`);

  const select = ms((i) => view.select(i % 50, 1), 50);
  console.log(`select          ${select.each.toFixed(2)} ms  ${bar(select.each)}`);

  const move = ms(() => view.moveSelection('down'), 50);
  console.log(`moveSelection   ${move.each.toFixed(2)} ms  ${bar(move.each)}`);

  const edit = ms((i) => view.setCell(i % 100, 4, `v${i}`), 30);
  console.log(`setCell         ${edit.each.toFixed(2)} ms  ${bar(edit.each)}`);

  const whole = ms(() => {
    view.moveSelection('down');
    return JSON.stringify(view.render());
  }, 20);
  console.log(`ROUND TRIP      ${whole.each.toFixed(2)} ms per arrow key  ${bar(whole.each)}`);
}

console.log('');
