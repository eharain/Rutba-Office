// Mailings → Create: Envelopes and Labels, as Word's own dialog has them.
//
// Envelopes adds the envelope in front of the letter as a section of its own
// — its own paper, turned the way it feeds, the return address in the corner
// and the delivery address in a frame where Word puts it — or prints it on
// its own. Labels makes a sheet of the same address, or one label of it, on
// any of the common Avery sheets, as a new document or straight to the
// printer. Start Mail Merge → Envelopes and → Labels open the same options
// and turn the document into the merge's envelope or sheet. The sizes and
// the sheets are `@rutba/ooxml/labels`' tables, the same ones the document
// is built from.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon } from '@rutba/office-ui';
import { ENVELOPE_SIZES, LABEL_PRODUCTS, labelProduct, describeLabel, addressLines } from '@rutba/ooxml/labels';
import { PrintDialog } from '../../print.js';

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Tahoma', 'Garamond', 'Consolas', 'Courier New'];
const SIZES = [8, 9, 10, 11, 12, 14, 16];
const RETURN_KEY = 'word.mailings.returnAddress';
const LABEL_KEY = 'word.mailings.label';
const ENVELOPE_KEY = 'word.mailings.envelope';

/** A card from the address book as the lines of an address. */
function contactLines(c) {
  const n = c.name || {};
  const name = [n.prefix, n.given || (!n.family ? n.full : ''), n.family].filter(Boolean).join(' ') || c.display || '';
  const a = (c.addresses || []).find((x) => x.pref) || (c.addresses || [])[0] || {};
  const street = String(a.street || '').split(/\r?\n/).filter(Boolean);
  const town = [a.city, a.region, a.postcode].filter(Boolean).join(' ');
  return [name, c.org, ...street, town, a.country].map((l) => String(l || '').trim()).filter(Boolean);
}

/** The address book, to pick one card from — Word's Address Book button. */
function AddressBook({ shell, onPick, onClose }) {
  const [query, setQuery] = useState('');
  const [list, setList] = useState(null);
  useEffect(() => {
    let live = true;
    shell.contacts.list({ query }).then((l) => live && setList(l || [])).catch(() => live && setList([]));
    return () => { live = false; };
  }, [shell, query]);
  return (
    <Dialog title="Select Name" width={460} onClose={onClose} actions={<Button label="Cancel" onClick={onClose} />}>
      <Input value={query} autoFocus placeholder="Search the address book" onChange={(e) => setQuery(e.target.value)} style={{ width: '100%', boxSizing: 'border-box' }} />
      <div className="wd-mm-pick wd-el-book" role="listbox" aria-label="Contacts">
        {list === null ? <span className="rw-hint" style={{ padding: 10 }}>Reading the address book…</span>
          : !list.length ? <span className="rw-hint" style={{ padding: 10 }}>No one here. People added in Contacts are listed.</span>
          : list.map((c) => {
            const lines = contactLines(c);
            return (
              <button key={c.id} type="button" onClick={() => onPick(lines)}>
                <Icon name="contacts" size={14} />
                <span><b>{lines[0] || c.display}</b>{lines.length > 1 ? <span className="rw-hint"> — {lines.slice(1).join(', ')}</span> : null}</span>
              </button>
            );
          })}
      </div>
    </Dialog>
  );
}

/** The envelope drawn to scale, the two addresses where they print. */
function EnvelopePreview({ size, delivery, returnAddress, delPt = 12, retPt = 10 }) {
  const s = ENVELOPE_SIZES.find((e) => e.id === size) || ENVELOPE_SIZES[0];
  // To scale, letters and all: the envelope fits the column, a big one
  // (C4) no taller than the room, and the words the size they print at.
  const scale = Math.min(330 / s.w, 190 / s.h);
  const fontOf = (pt) => Math.max(4.5, pt * 20 * scale);
  const w = Math.round(s.w * scale);
  const h = Math.round(s.h * scale);
  const frameX = Math.round(((s.w - 7920) / 2 + 2880) * scale);
  const frameY = Math.round((s.h - 1980) * scale);
  return (
    <div className="wd-el-envelope" style={{ width: w, height: h }} aria-label="Envelope preview">
      {returnAddress?.length ? <div className="wd-el-ret" style={{ left: Math.round(576 * scale), top: Math.round(360 * scale), fontSize: fontOf(retPt) }}>{returnAddress.map((l, i) => <div key={i}>{l}</div>)}</div> : null}
      <div className="wd-el-del" style={{ left: frameX, top: frameY, fontSize: fontOf(delPt) }}>{delivery.length ? delivery.map((l, i) => <div key={i}>{l}</div>) : <span className="wd-el-empty">Delivery address</span>}</div>
      <div className="wd-el-stamp" />
    </div>
  );
}

/** The sheet drawn to scale: its labels, the one to print picked out. */
function SheetPreview({ product, mode, row, col }) {
  const p = labelProduct(product);
  const scale = 118 / p.page.w;
  return (
    <div className="wd-el-sheet" style={{ width: Math.round(p.page.w * scale), height: Math.round(p.page.h * scale) }} aria-label="Sheet preview">
      {Array.from({ length: p.rows }, (_, r) => Array.from({ length: p.cols }, (__, c) => (
        <div
          key={`${r}:${c}`}
          className={`wd-el-label${mode === 'full' || (r === row - 1 && c === col - 1) ? ' on' : ''}`}
          style={{ left: Math.round((p.side + c * p.pitchX) * scale), top: Math.round((p.top + r * p.pitchY) * scale), width: Math.max(2, Math.round(p.w * scale) - 1), height: Math.max(2, Math.round(p.h * scale) - 1) }}
        />
      )))}
    </div>
  );
}

/** Font and size, side by side. */
function FontPick({ label, value, onChange }) {
  return (
    <Field label={label}>
      <div className="wd-el-font">
        <Select value={value.name} onChange={(e) => onChange({ ...value, name: e.target.value })}>{FONTS.map((f) => <option key={f} value={f}>{f}</option>)}</Select>
        <Select value={String(value.sizePt)} onChange={(e) => onChange({ ...value, sizePt: Number(e.target.value) })}>{SIZES.map((s) => <option key={s} value={String(s)}>{s} pt</option>)}</Select>
      </div>
    </Field>
  );
}

/* ── Envelopes ───────────────────────────────────────────────────────────── */

export function EnvelopesDialog({ shell, model, selection, onClose, onAdd, onPrint, mergeMode = false }) {
  const had = model?.envelope || null;
  const [delivery, setDelivery] = useState(() => (had?.delivery?.length ? had.delivery.join('\n') : selection || ''));
  const [ret, setRet] = useState(() => (had?.returnAddress?.length ? had.returnAddress.join('\n') : ''));
  const [omit, setOmit] = useState(false);
  const [size, setSize] = useState(had?.size || 'DL');
  const [delFont, setDelFont] = useState({ name: 'Calibri', sizePt: 12 });
  const [retFont, setRetFont] = useState({ name: 'Calibri', sizePt: 10 });
  const [book, setBook] = useState(null);
  const touched = useRef(false);
  // The return address Word remembers, from the last envelope.
  useEffect(() => {
    shell.store.get({ key: RETURN_KEY, fallback: '' }).then((v) => { if (!touched.current && !had?.returnAddress?.length && v) setRet(v); }).catch(() => {});
    shell.store.get({ key: ENVELOPE_KEY, fallback: null }).then((v) => {
      if (!v || touched.current) return;
      if (!had?.size && v.size) setSize(v.size);
      if (v.delFont) setDelFont(v.delFont);
      if (v.retFont) setRetFont(v.retFont);
    }).catch(() => {});
  }, [shell, had]);
  const spec = () => ({
    size,
    delivery: addressLines(delivery),
    returnAddress: omit ? null : addressLines(ret),
    deliveryFont: delFont,
    returnFont: retFont,
  });
  const remember = () => {
    if (!omit) shell.store.set({ key: RETURN_KEY, value: ret }).catch(() => {});
    shell.store.set({ key: ENVELOPE_KEY, value: { size, delFont, retFont } }).catch(() => {});
  };
  return (
    <>
      <Dialog
        title={mergeMode ? 'Envelope Options' : 'Envelopes'}
        width={780}
        onClose={onClose}
        actions={
          <>
            {mergeMode ? null : <Button icon="print" label="Print" disabled={!addressLines(delivery).length} onClick={() => { remember(); onPrint(spec()); }} />}
            <span style={{ flex: 1 }} />
            <Button label="Cancel" onClick={onClose} />
            <Button primary label={mergeMode ? 'OK' : had ? 'Change Document' : 'Add to Document'} disabled={!mergeMode && !addressLines(delivery).length} onClick={() => { remember(); onAdd(spec()); }} />
          </>
        }
      >
        <div className="wd-el-cols">
          <div className="wd-el-col">
            {mergeMode ? (
              <p className="wd-mm-lead">The document becomes the envelope. Put an Address Block in the delivery address frame to merge each recipient's address.</p>
            ) : (
              <Field label="Delivery address:">
                <div className="wd-el-addr">
                  <textarea className="rw-input wd-el-text wd-el-delivery" rows={5} value={delivery} autoFocus onChange={(e) => { touched.current = true; setDelivery(e.target.value); }} />
                  <Button icon="contacts" title="Address Book — insert an address from Contacts" onClick={() => setBook('delivery')} />
                </div>
              </Field>
            )}
            <Field label="Return address:">
              <div className="wd-el-addr">
                <textarea className="rw-input wd-el-text wd-el-return" rows={4} value={ret} disabled={omit} onChange={(e) => { touched.current = true; setRet(e.target.value); }} />
                <Button icon="contacts" title="Address Book — insert an address from Contacts" disabled={omit} onClick={() => setBook('return')} />
              </div>
            </Field>
            <label className="wd-mm-check"><input type="checkbox" checked={omit} onChange={(e) => setOmit(e.target.checked)} /> Omit</label>
          </div>
          <div className="wd-el-col">
            <Field label="Envelope size:">
              <Select className="rw-select wd-el-size" value={size} onChange={(e) => { touched.current = true; setSize(e.target.value); }} style={{ width: '100%' }}>
                {ENVELOPE_SIZES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}
              </Select>
            </Field>
            <div className="wd-el-fonts">
              <FontPick label="Delivery address font:" value={delFont} onChange={setDelFont} />
              <FontPick label="Return address font:" value={retFont} onChange={setRetFont} />
            </div>
            <div className="wd-mm-section">Preview</div>
            <EnvelopePreview size={size} delivery={mergeMode ? ['«AddressBlock»'] : addressLines(delivery)} returnAddress={omit ? [] : addressLines(ret)} delPt={delFont.sizePt} retPt={retFont.sizePt} />
            <p className="rw-hint">{had && !mergeMode ? 'Change Document replaces the envelope already in front of this document.' : 'Add to Document puts the envelope in front of the letter, on a page of its own size.'}</p>
          </div>
        </div>
      </Dialog>
      {book ? (
        <AddressBook
          shell={shell}
          onClose={() => setBook(null)}
          onPick={(lines) => { touched.current = true; if (book === 'delivery') setDelivery(lines.join('\n')); else setRet(lines.join('\n')); setBook(null); }}
        />
      ) : null}
    </>
  );
}

/* ── Labels ──────────────────────────────────────────────────────────────── */

/** Label Options: the vendor, the product, and what the product is. */
export function LabelOptionsDialog({ product, onClose, onPick, title = 'Label Options' }) {
  const vendors = [...new Set(LABEL_PRODUCTS.map((p) => p.vendor))];
  const [vendor, setVendor] = useState(labelProduct(product).vendor);
  const [picked, setPicked] = useState(labelProduct(product).id);
  const list = LABEL_PRODUCTS.filter((p) => p.vendor === vendor);
  const p = labelProduct(picked);
  return (
    <Dialog
      title={title}
      width={600}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="OK" onClick={() => onPick(picked)} /></>}
    >
      <Field label="Label vendors:">
        <Select className="rw-select wd-el-vendor" value={vendor} onChange={(e) => { setVendor(e.target.value); setPicked(LABEL_PRODUCTS.find((x) => x.vendor === e.target.value).id); }} style={{ width: '100%' }}>
          {vendors.map((v) => <option key={v} value={v}>{v}</option>)}
        </Select>
      </Field>
      <div className="wd-el-cols wd-el-options">
        <div className="wd-el-col">
          <div className="wd-mm-section">Product number:</div>
          <div className="wd-mm-pick wd-el-products" role="listbox" aria-label="Product number">
            {list.map((x) => (
              <button key={x.id} type="button" role="option" aria-selected={x.id === picked} className={x.id === picked ? 'on' : ''} data-product={x.id} onClick={() => setPicked(x.id)} onDoubleClick={() => onPick(x.id)}>
                {x.name} {x.kind} Labels
              </button>
            ))}
          </div>
        </div>
        <div className="wd-el-col">
          <div className="wd-mm-section">Label information</div>
          <dl className="about-list wd-el-info">
            <dt>Type</dt><dd>{p.kind} Labels</dd>
            <dt>Height</dt><dd>{Math.round(p.h * 10) / 10} mm</dd>
            <dt>Width</dt><dd>{Math.round(p.w * 10) / 10} mm</dd>
            <dt>Page size</dt><dd>{p.page.name} — {p.page.w} × {p.page.h} mm</dd>
            <dt>Labels</dt><dd>{p.cols} across, {p.rows} down</dd>
          </dl>
          <SheetPreview product={picked} mode="full" row={1} col={1} />
        </div>
      </div>
    </Dialog>
  );
}

export function LabelsDialog({ shell, selection, onClose, onNewDocument, onPrint }) {
  const [text, setText] = useState(selection || '');
  const [useReturn, setUseReturn] = useState(false);
  const [returnText, setReturnText] = useState('');
  const [mode, setMode] = useState('full');
  const [row, setRow] = useState('1');
  const [col, setCol] = useState('1');
  const [product, setProduct] = useState('L7160');
  const [options, setOptions] = useState(false);
  const [book, setBook] = useState(false);
  useEffect(() => {
    shell.store.get({ key: RETURN_KEY, fallback: '' }).then((v) => setReturnText(v || '')).catch(() => {});
    shell.store.get({ key: LABEL_KEY, fallback: null }).then((v) => { if (v && LABEL_PRODUCTS.some((p) => p.id === v)) setProduct(v); }).catch(() => {});
  }, [shell]);
  const p = labelProduct(product);
  const lines = addressLines(useReturn ? returnText : text);
  const spec = () => ({ product, lines, mode, row: Math.min(p.rows, Math.max(1, Number(row) || 1)), col: Math.min(p.cols, Math.max(1, Number(col) || 1)) });
  return (
    <>
      <Dialog
        title="Labels"
        width={760}
        onClose={onClose}
        actions={
          <>
            <Button icon="print" label="Print" disabled={!lines.length} onClick={() => onPrint(spec())} />
            <span style={{ flex: 1 }} />
            <Button label="Cancel" onClick={onClose} />
            <Button primary label="New Document" disabled={!lines.length} onClick={() => onNewDocument(spec())} />
          </>
        }
      >
        <div className="wd-el-cols">
          <div className="wd-el-col">
            <Field label="Address:">
              <div className="wd-el-addr">
                <textarea className="rw-input wd-el-text wd-el-labeltext" rows={5} value={useReturn ? returnText : text} disabled={useReturn} autoFocus onChange={(e) => setText(e.target.value)} />
                <Button icon="contacts" title="Address Book — insert an address from Contacts" disabled={useReturn} onClick={() => setBook(true)} />
              </div>
            </Field>
            <label className="wd-mm-check"><input type="checkbox" checked={useReturn} disabled={!returnText} onChange={(e) => setUseReturn(e.target.checked)} /> Use return address</label>
            <div className="wd-mm-section">Print</div>
            <div className="wd-mm-radios">
              <label><input type="radio" name="wd-el-mode" checked={mode === 'full'} onChange={() => setMode('full')} /> Full page of the same label</label>
              <label className="wd-mm-inline"><input type="radio" name="wd-el-mode" checked={mode === 'single'} onChange={() => setMode('single')} /> Single label — Row:
                <Input className="rw-input wd-mm-num" value={row} disabled={mode !== 'single'} onChange={(e) => setRow(e.target.value.replace(/[^\d]/g, ''))} aria-label="Row" />
                Column:
                <Input className="rw-input wd-mm-num" value={col} disabled={mode !== 'single'} onChange={(e) => setCol(e.target.value.replace(/[^\d]/g, ''))} aria-label="Column" />
              </label>
            </div>
          </div>
          <div className="wd-el-col">
            <div className="wd-mm-section">Label</div>
            <div className="wd-el-product">
              <div>
                <div className="wd-el-product-name">{p.vendor}, {p.name}</div>
                <div className="rw-hint">{describeLabel(p)}</div>
              </div>
              <Button label="Options…" onClick={() => setOptions(true)} />
            </div>
            <SheetPreview product={product} mode={mode} row={Number(row) || 1} col={Number(col) || 1} />
          </div>
        </div>
      </Dialog>
      {options ? <LabelOptionsDialog product={product} onClose={() => setOptions(false)} onPick={(id) => { setProduct(id); shell.store.set({ key: LABEL_KEY, value: id }).catch(() => {}); setOptions(false); }} /> : null}
      {book ? <AddressBook shell={shell} onClose={() => setBook(false)} onPick={(l) => { setText(l.join('\n')); setBook(false); }} /> : null}
    </>
  );
}

/* ── the verbs, for the Word window ──────────────────────────────────────── */

/**
 * Envelopes, Labels, and the two Start Mail Merge choices that lay the
 * document out as an envelope or a sheet — for the Mailings tab. `node`
 * is the dialogs; `extra` joins the mail merge verbs.
 */
export function useEnvelopesLabels({ shell, doc, model, apply, toast, selectionText }) {
  const [dialog, setDialog] = useState(null);
  const [printing, setPrinting] = useState(null);
  const call = (action, args) => shell.doc.mailMerge({ id: doc.id, action, ...args });
  const close = () => setDialog(null);

  const extra = useMemo(() => ({
    envelopes: () => setDialog({ name: 'envelopes', selection: selectionText?.() || '' }),
    labels: () => setDialog({ name: 'labels', selection: selectionText?.() || '' }),
    startLayout: (type) => setDialog({ name: type === 'envelopes' ? 'mergeEnvelope' : 'mergeLabels' }),
    updateLabels: async () => {
      const next = await apply({ op: 'updateLabels' });
      if (next) toast(`Updated ${next.opResult || 0} label${next.opResult === 1 ? '' : 's'} from the first one.`, { tone: 'good' });
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [apply, toast, selectionText]);

  // A command naming a label sheet or an envelope — the OS, or a check run —
  // goes through the same doors as the dialogs' buttons.
  const actRef = useRef(null);
  actRef.current = { newLabels: (spec) => newLabels(spec), addEnvelope: (spec) => apply({ op: 'addEnvelope', spec }) };
  useEffect(() => {
    if (!shell?.on) return undefined;
    return shell.on('app:command', ({ command, args }) => {
      if (command === 'mailings.newLabels' && args) actRef.current.newLabels(args);
      if (command === 'mailings.addEnvelope' && args) actRef.current.addEnvelope(args);
    });
  }, [shell]);

  async function newLabels(spec) {
    try {
      const made = await call('makeLabels', { spec });
      await shell.win.create({ app: 'word', query: { session: made.id } });
      toast(`${made.name}: a sheet of ${labelProduct(spec.product).name} labels.`, { tone: 'good' });
      close();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }
  async function printMade(action, spec) {
    try {
      const made = await call(action, { spec });
      close();
      setPrinting(made);
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }

  const node = (
    <>
      {dialog?.name === 'envelopes' ? (
        <EnvelopesDialog
          shell={shell}
          model={model}
          selection={dialog.selection}
          onClose={close}
          onAdd={async (spec) => { close(); const next = await apply({ op: 'addEnvelope', spec }); if (next) toast('The envelope is the first page of the document.', { tone: 'good' }); }}
          onPrint={(spec) => printMade('makeEnvelope', spec)}
        />
      ) : null}
      {dialog?.name === 'mergeEnvelope' ? (
        <EnvelopesDialog
          shell={shell}
          model={null}
          mergeMode
          onClose={close}
          onAdd={async (spec) => { close(); await apply({ op: 'setEnvelopeDocument', spec }); }}
        />
      ) : null}
      {dialog?.name === 'labels' ? (
        <LabelsDialog shell={shell} selection={dialog.selection} onClose={close} onNewDocument={newLabels} onPrint={(spec) => printMade('makeLabels', spec)} />
      ) : null}
      {dialog?.name === 'mergeLabels' ? (
        <LabelOptionsDialog
          title="Label Options"
          product="L7160"
          onClose={close}
          onPick={async (id) => {
            close();
            shell.store.set({ key: LABEL_KEY, value: id }).catch(() => {});
            const next = await apply({ op: 'setLabelSheet', spec: { product: id, mode: 'merge' } });
            if (next) toast('A sheet of labels, each after the first starting with «Next Record». Put an Address Block in the first, then Update Labels.', { ms: 7000 });
          }}
        />
      ) : null}
      {printing ? (
        <PrintDialog
          shell={shell}
          doc={{ id: printing.id, name: printing.name }}
          kind="doc"
          onClose={() => { shell.doc.close({ id: printing.id }).catch(() => {}); setPrinting(null); }}
          onSaveAs={null}
        />
      ) : null}
    </>
  );
  return { extra, node };
}

export const ENVELOPES_CSS = `
.wd-el-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 22px; }
.wd-el-col { min-width: 0; display: flex; flex-direction: column; gap: 10px; }
.wd-el-addr { display: flex; gap: 6px; align-items: flex-start; }
.wd-el-text.rw-input { flex: 1; min-width: 0; resize: vertical; font: inherit; line-height: 1.4; box-sizing: border-box; }
.wd-el-fonts { display: grid; grid-template-columns: 1fr; gap: 8px; }
.wd-el-font { display: grid; grid-template-columns: 1fr 86px; gap: 6px; }
.wd-el-envelope { position: relative; background: #fbfaf6; border: 1px solid var(--line-strong); border-radius: 3px; box-shadow: 0 1px 3px rgba(0,0,0,.08); font-family: Calibri, 'Segoe UI', sans-serif; color: #222; overflow: hidden; }
.wd-el-ret { position: absolute; line-height: 1.2; white-space: nowrap; }
.wd-el-del { position: absolute; line-height: 1.2; white-space: nowrap; }
.wd-el-empty { color: #999; font-style: italic; }
.wd-el-stamp { position: absolute; right: 8px; top: 8px; width: 22px; height: 26px; border: 1px dashed #bbb; border-radius: 2px; }
.wd-el-sheet { position: relative; background: #fff; border: 1px solid var(--line-strong); box-shadow: 0 1px 3px rgba(0,0,0,.08); align-self: flex-start; }
.wd-el-label { position: absolute; border: 1px solid #d5d9e0; border-radius: 2px; background: #fff; }
.wd-el-label.on { background: color-mix(in srgb, var(--accent) 22%, white); border-color: var(--accent-line); }
.wd-el-product { display: flex; align-items: center; gap: 10px; justify-content: space-between; border: 1px solid var(--line); border-radius: var(--r-2); padding: 8px 10px; background: var(--surface-2); }
.wd-el-product-name { font-weight: 600; font-size: 12.5px; }
.wd-el-product .rw-btn { border: 1px solid var(--line); background: var(--surface); }
.wd-el-products { max-height: 230px; }
.wd-el-info { margin: 0; }
.wd-el-book { margin-top: 8px; max-height: 300px; }
.wd-el-book button { align-items: flex-start; }
.rw-dialog[aria-label="Envelopes"], .rw-dialog[aria-label="Labels"], .rw-dialog[aria-label="Envelope Options"] { max-width: min(94vw, 820px); }
`;

let installed = false;
export function installEnvelopeStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-envelopes-css';
  style.textContent = ENVELOPES_CSS;
  document.head.appendChild(style);
}
