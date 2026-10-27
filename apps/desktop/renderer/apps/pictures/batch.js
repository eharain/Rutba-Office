// Batch from Pictures: many pictures resized, converted, renamed by a
// pattern, turned upright by their EXIF orientation and their metadata
// stripped, in one go. The planning is @rutba/imaging's batch.js; the
// decoding and encoding happen in the image worker, so neither the main
// process nor the window's own thread does the work; the files are only read
// and written through the shell. Nothing already on disk is replaced without
// asking: a clash is found before the first file is written, and the person
// chooses Keep both, Replace or Cancel.
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, Button, Select, Input, Progress, basename, t, tn } from '@rutba/office-ui';
import { probeImage } from '@rutba/imaging/probe';
import { readExif } from '@rutba/imaging/exif';
import { planBatch, keepBoth, nameFromPattern, sizeFor, folderOf, formatOf, FORMATS } from '@rutba/imaging/batch';
import { convertOff } from '../../imaging/work.js';

/** What the worker can decode: the pictures a batch takes. */
export const BATCH_EXTS = ['.jpg', '.jpeg', '.jpe', '.png', '.webp', '.bmp', '.gif', '.avif'];
export const canBatch = (file) => BATCH_EXTS.includes(String(file).toLowerCase().match(/\.[a-z0-9]+$/)?.[0] || '');

const SIZE_MODES = [
  { id: 'none', label: t('As they are') },
  { id: 'percent', label: t('A percentage') },
  { id: 'fit', label: t('To fit inside') },
  { id: 'long', label: t('Longest side') },
];

// What a name pattern can hold, each pressed to add it.
const TOKENS = [
  ['{name}', t('Original name')],
  ['{n}', t('Number')],
  ['{n:3}', t('Number as 001')],
  ['{date}', t('File date')],
  ['{w}', t('Width')],
  ['{h}', t('Height')],
];

/** The size a picture will be read at: its stored size, turned if its EXIF says so and it is to be turned. */
async function sizeOf(shell, file, upright) {
  try {
    const { bytes } = await shell.fs.readHead({ path: file, bytes: 256 * 1024 });
    const info = probeImage(bytes) || {};
    const turn = upright && [5, 6, 7, 8].includes(readExif(bytes).orientation);
    return turn ? { width: info.height || 0, height: info.width || 0 } : { width: info.width || 0, height: info.height || 0 };
  } catch {
    return { width: 0, height: 0 };
  }
}

/**
 * The choices of the last batch, remembered on this computer for the next:
 * the size, format, quality, name pattern, upright and strip. Where the
 * pictures go is not: beside the originals, each time, unless chosen.
 */
const BATCH_KEY = 'pictures.batch';
const BATCH_DEFAULTS = { mode: 'fit', percent: 50, box: { width: 1920, height: 1080 }, long: 1600, format: 'jpeg', quality: 85, pattern: '{name} (edited)', upright: true, strip: 'all' };

/**
 * What of the camera's details each picture keeps: none of them (the
 * default, as before), all but the place, or all. Kept in a JPEG, PNG or
 * WebP; the worker sets the orientation upright where the pixels are.
 */
export const METADATA = [
  { id: 'all', role: 'batch-strip', label: t('Strip all: no camera, date taken or place') },
  { id: 'location', role: 'batch-location', label: t('Remove location only: keep the camera and date taken') },
  { id: 'none', role: 'batch-keep', label: t('Keep all, the place too') },
];
export function rememberedBatch() {
  try {
    const kept = JSON.parse(localStorage.getItem(BATCH_KEY) || '{}');
    const out = { ...BATCH_DEFAULTS };
    for (const k of Object.keys(BATCH_DEFAULTS)) {
      const v = kept?.[k];
      if (v == null) continue;
      if (k === 'mode' && !SIZE_MODES.some((m) => m.id === v)) continue;
      if (k === 'format' && v !== 'same' && !FORMATS[v]) continue;
      if (k === 'box' && !(v && v.width > 0 && v.height > 0)) continue;
      if (['percent', 'long', 'quality'].includes(k) && !(Number(v) > 0)) continue;
      if (k === 'pattern' && !String(v).trim()) continue;
      if (k === 'upright' && typeof v !== 'boolean') continue;
      if (k === 'strip') {
        // Kept as a yes or no before Remove location came.
        const was = v === true ? 'all' : v === false ? 'none' : v;
        if (METADATA.some((m) => m.id === was)) out.strip = was;
        continue;
      }
      out[k] = v;
    }
    return out;
  } catch {
    return { ...BATCH_DEFAULTS };
  }
}

export function BatchDialog({ shell, files, folder, onClose, onDone }) {
  const pictures = useMemo(() => files.filter((f) => !f.dir && canBatch(f.path)), [files]);
  const [chosen, setChosen] = useState(() => new Set(pictures.map((f) => f.path)));
  const last = useMemo(rememberedBatch, []);
  const [mode, setMode] = useState(last.mode);
  const [percent, setPercent] = useState(last.percent);
  const [box, setBox] = useState(last.box);
  const [long, setLong] = useState(last.long);
  const [format, setFormat] = useState(last.format);
  const [quality, setQuality] = useState(last.quality);
  const [pattern, setPattern] = useState(last.pattern);
  const [upright, setUpright] = useState(last.upright);
  const [strip, setStrip] = useState(last.strip);
  const [out, setOut] = useState(null); // null: beside the originals
  const [run, setRun] = useState(null); // { done, total, name, written, failed, stopped }
  const stop = useRef(false);

  const resize = mode === 'percent' ? { mode, percent } : mode === 'fit' ? { mode, ...box } : mode === 'long' ? { mode, size: long } : { mode: 'none' };
  const picked = pictures.filter((f) => chosen.has(f.path));
  const example = picked[0] || pictures[0];
  const exampleName = example
    ? `${nameFromPattern(pattern, { name: basename(example.path).replace(/\.[^.]+$/, ''), n: 1, date: example.mtime, width: 1600, height: 1200 })}.${FORMATS[format === 'same' ? formatOf(example.path) || 'png' : format].ext}`
    : '';

  const chooseFolder = async () => {
    const [dir] = await shell.dialog.open({ title: t('Save the pictures in'), directory: true, defaultPath: out || folder });
    if (dir) setOut(dir);
  };

  // Each choice kept as it is made, for the next batch.
  useEffect(() => {
    try { localStorage.setItem(BATCH_KEY, JSON.stringify({ mode, percent, box, long, format, quality, pattern, upright, strip })); } catch { /* no storage, no memory */ }
  }, [mode, percent, box, long, format, quality, pattern, upright, strip]);

  const start = async () => {
    stop.current = false;
    setRun({ done: 0, total: picked.length, name: '', written: [], failed: [], stopped: false });
    // The sizes first, for names that carry them; then what is on disk where the files will go.
    const items = [];
    for (const f of picked) items.push({ path: f.path, date: f.mtime, size: await sizeOf(shell, f.path, upright) });
    let plan = planBatch(items, { format, pattern, folder: out, resize });
    const there = new Set();
    for (const dir of new Set(plan.items.map((it) => folderOf(it.target)))) {
      for (const row of await shell.fs.list({ path: dir }).catch(() => [])) there.add(String(row.path).toLowerCase());
    }
    const exists = (p) => there.has(String(p).toLowerCase());
    plan = planBatch(items, { format, pattern, folder: out, resize, exists });
    if (plan.clashes.length) {
      const originals = plan.clashes.some((c) => c.original);
      const { response } = await shell.dialog.message({
        type: 'question',
        message: tn(plan.clashes.length, '{count} of these files is already there.', '{count} of these files are already there.'),
        detail: `${plan.clashes.slice(0, 6).map((c) => basename(c.target)).join('\n')}${plan.clashes.length > 6 ? '\n…' : ''}${originals ? `\n\n${t('Replace would write over the original pictures themselves.')}` : ''}`,
        buttons: [t('Keep both'), t('Replace'), t('Cancel')],
        defaultId: 0,
        cancelId: 2,
      });
      if (response === 2 || response === undefined) { setRun(null); return; }
      if (response === 0) plan = keepBoth(plan, exists);
    }
    if (out) await shell.fs.mkdirp({ path: out }).catch(() => {});
    const written = [];
    const failed = [];
    for (let i = 0; i < plan.items.length; i++) {
      if (stop.current) break;
      const it = plan.items[i];
      setRun((r) => ({ ...r, done: i, name: basename(it.source) }));
      try {
        const { bytes } = await shell.fs.read({ path: it.source });
        const done = await convertOff(bytes, { format: it.format, quality: quality / 100, resize, upright, strip });
        await shell.fs.write({ path: it.target, bytes: done.bytes });
        written.push({ ...it, width: done.width, height: done.height, size: done.bytes.length });
      } catch {
        failed.push(it);
      }
    }
    setRun({ done: written.length + failed.length, total: plan.items.length, name: '', written, failed, stopped: stop.current, finished: true });
    onDone?.(written);
  };

  const finished = run?.finished;
  const toggle = (p) => setChosen((s) => { const n = new Set(s); if (n.has(p)) n.delete(p); else n.add(p); return n; });
  const num = (v, set) => <Input type="number" min={1} value={v} onChange={(e) => set(Number(e.target.value) || 1)} style={{ width: 76 }} />;

  return (
    <Dialog
      title={t('Batch edit')}
      width={560}
      className="im-batch"
      onClose={run && !finished ? undefined : onClose}
      actions={
        finished ? (
          <>
            {run.written.length ? <Button label={t('Show in folder')} onClick={() => shell.shell.showInFolder({ path: run.written[0].target })} /> : null}
            <Button primary label={t('Close')} onClick={onClose} data-role="batch-close" />
          </>
        ) : run ? (
          <Button label={t('Stop')} onClick={() => { stop.current = true; }} />
        ) : (
          <>
            <Button label={t('Cancel')} onClick={onClose} />
            <Button primary label={tn(picked.length, 'Convert {count} picture', 'Convert {count} pictures')} disabled={!picked.length} onClick={start} data-role="batch-run" />
          </>
        )
      }
    >
      {run ? (
        <div className="im-batch-run" data-state={finished ? 'done' : 'running'}>
          <Progress value={finished ? 1 : run.done / Math.max(1, run.total)} />
          {finished ? (
            <>
              <div className="im-batch-sum" data-written={run.written.length} data-failed={run.failed.length}>
                {run.stopped ? t('Stopped.') : null}{' '}
                {tn(run.written.length, '{count} picture written', '{count} pictures written')}
                {run.failed.length ? `, ${tn(run.failed.length, '{count} could not be converted', '{count} could not be converted')}` : ''}
              </div>
              <ul className="im-batch-out">
                {run.written.map((w) => (
                  <li key={w.target}><span>{basename(w.target)}</span><span>{`${w.width} × ${w.height}`}</span></li>
                ))}
                {run.failed.map((w) => (
                  <li key={w.source} className="bad"><span>{basename(w.source)}</span><span>{t('Not converted')}</span></li>
                ))}
              </ul>
            </>
          ) : (
            <div className="rw-hint">{t('Converting {done} of {total}: {name}', { done: run.done + 1, total: run.total, name: run.name })}</div>
          )}
        </div>
      ) : (
        <div className="im-batch-form">
          <section className="im-batch-files">
            <div className="im-batch-head">
              <strong>{tn(picked.length, '{count} picture chosen', '{count} pictures chosen')}</strong>
              <span className="spacer" />
              <Button label={t('All')} onClick={() => setChosen(new Set(pictures.map((f) => f.path)))} />
              <Button label={t('None')} onClick={() => setChosen(new Set())} />
            </div>
            <div className="im-batch-list">
              {pictures.map((f) => (
                <label key={f.path} className="im-batch-file" data-file={basename(f.path)}>
                  <input type="checkbox" checked={chosen.has(f.path)} onChange={() => toggle(f.path)} />
                  <span>{basename(f.path)}</span>
                </label>
              ))}
              {pictures.length ? null : <div className="rw-hint">{t('This folder has no pictures a batch can convert.')}</div>}
            </div>
          </section>

          <div className="im-batch-grid">
            <span className="im-batch-label">{t('Size')}</span>
            <div className="im-batch-row">
              <Select value={mode} onChange={(e) => setMode(e.target.value)} data-role="batch-size" style={{ width: 140 }}>
                {SIZE_MODES.map((m) => <option key={m.id} value={m.id}>{m.label}</option>)}
              </Select>
              {mode === 'percent' ? <>{num(percent, setPercent)}<span>%</span></> : null}
              {mode === 'fit' ? (
                <>
                  {num(box.width, (v) => setBox((b) => ({ ...b, width: v })))}
                  <span>×</span>
                  {num(box.height, (v) => setBox((b) => ({ ...b, height: v })))}
                  <span>{t('px')}</span>
                </>
              ) : null}
              {mode === 'long' ? <>{num(long, setLong)}<span>{t('px')}</span></> : null}
            </div>

            <span className="im-batch-label">{t('Format')}</span>
            <div className="im-batch-row">
              <Select value={format} onChange={(e) => setFormat(e.target.value)} data-role="batch-format" style={{ width: 140 }}>
                <option value="same">{t('Same as the original')}</option>
                <option value="jpeg">JPEG</option>{/* words-ok: a format's name */}
                <option value="png">PNG</option>{/* words-ok: a format's name */}
                <option value="webp">WebP</option>{/* words-ok: a format's name */}
              </Select>
              {format === 'jpeg' || format === 'webp' || format === 'same' ? (
                <label className="im-batch-quality">
                  <span>{t('Quality')}</span>
                  <input type="range" min={10} max={100} value={quality} onChange={(e) => setQuality(Number(e.target.value))} data-role="batch-quality" />
                  <span className="im-ps-value">{quality}%</span>
                </label>
              ) : null}
            </div>

            <span className="im-batch-label">{t('Name')}</span>
            <div className="im-batch-col">
              <Input value={pattern} onChange={(e) => setPattern(e.target.value)} data-role="batch-pattern" style={{ width: '100%' }} />
              <div className="im-batch-tokens">
                {TOKENS.map(([token, label]) => (
                  <button key={token} type="button" className="im-batch-token" data-tip={label} onClick={() => setPattern((p) => `${p}${token}`)}>
                    <code>{token}</code>
                    <span>{label}</span>
                  </button>
                ))}
              </div>
              {exampleName ? <div className="rw-hint" data-role="batch-example">{t('The first will be {example}', { example: exampleName })}</div> : null}
            </div>

            <span className="im-batch-label">{t('Save in')}</span>
            <div className="im-batch-row">
              <Select value={out ? 'folder' : 'beside'} onChange={(e) => (e.target.value === 'beside' ? setOut(null) : chooseFolder())} data-role="batch-where" style={{ width: 200 }}>
                <option value="beside">{t('Next to the originals')}</option>
                <option value="folder">{out ? basename(out) : t('A folder of my choosing…')}</option>
              </Select>
              {out ? <Button label={t('Change…')} onClick={chooseFolder} /> : null}
            </div>

            <span className="im-batch-label">{t('Also')}</span>
            <div className="im-batch-col">
              <label className="im-check">
                <input type="checkbox" checked={upright} onChange={(e) => setUpright(e.target.checked)} data-role="batch-upright" />
                <span>{t('Turn each picture upright by its EXIF orientation')}</span>
              </label>
            </div>

            <span className="im-batch-label">{t('Details')}</span>
            <div className="im-batch-col" role="radiogroup" aria-label={t('The details each picture keeps')}>
              {METADATA.map((m) => (
                <label key={m.id} className="im-check">
                  <input type="radio" name="batch-metadata" value={m.id} checked={strip === m.id} onChange={() => setStrip(m.id)} data-role={m.role} />
                  <span>{m.label}</span>
                </label>
              ))}
            </div>
          </div>
          <div className="rw-hint im-batch-foot">{t('The originals are never changed. A file that is already there is never replaced without asking.')}</div>
        </div>
      )}
      <style>{BATCH_CSS}</style>
    </Dialog>
  );
}

const BATCH_CSS = `
.im-batch-form { display: flex; flex-direction: column; gap: 14px; }
.im-batch-files { border: 1px solid var(--line); border-radius: 6px; overflow: hidden; }
.im-batch-head { display: flex; align-items: center; gap: 4px; padding: 6px 8px; background: var(--sunken); font-size: 12.5px; }
.im-batch-head .spacer { flex: 1; }
.im-batch-list { max-height: 132px; overflow: auto; padding: 4px 8px; display: grid; grid-template-columns: 1fr 1fr; gap: 0 12px; }
.im-batch-file { display: flex; align-items: center; gap: 7px; font-size: 12.5px; padding: 3px 0; min-width: 0; }
.im-batch-file span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.im-batch-grid { display: grid; grid-template-columns: 76px 1fr; gap: 12px 12px; align-items: start; }
.im-batch-label { font-size: 12.5px; color: var(--ink-2); padding-top: 6px; }
.im-batch-row { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); flex-wrap: wrap; }
.im-batch-col { display: flex; flex-direction: column; gap: 6px; }
.im-batch-quality { display: flex; align-items: center; gap: 8px; font-size: 12.5px; }
.im-batch-quality input { width: 120px; accent-color: var(--accent); }
.im-batch input[type=checkbox], .im-batch input[type=radio] { accent-color: var(--accent); }
.im-batch .im-check { display: flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink); }
.im-batch-tokens { display: flex; flex-wrap: wrap; gap: 4px; }
.im-batch-token { display: inline-flex; align-items: center; gap: 5px; border: 1px solid var(--line); background: var(--bg, #fff);
  border-radius: 4px; padding: 2px 6px; font: inherit; font-size: 11.5px; color: var(--ink-2); cursor: pointer; }
.im-batch-token:hover { border-color: var(--accent); color: var(--ink); }
.im-batch-token code { font-family: Consolas, ui-monospace, monospace; color: var(--accent-text); font-size: 11px; }
.im-batch-foot { border-top: 1px solid var(--line-soft); padding-top: 10px; }
.im-batch-run { display: flex; flex-direction: column; gap: 12px; min-height: 120px; }
.im-batch-sum { font-size: 13px; font-weight: 600; }
.im-batch-out { list-style: none; margin: 0; padding: 0; max-height: 220px; overflow: auto; border: 1px solid var(--line-soft); border-radius: 6px; }
.im-batch-out li { display: flex; justify-content: space-between; gap: 12px; padding: 5px 10px; font-size: 12.5px; border-bottom: 1px solid var(--line-soft); }
.im-batch-out li:last-child { border-bottom: 0; }
.im-batch-out li span:last-child { color: var(--ink-3); font-variant-numeric: tabular-nums; }
.im-batch-out li.bad span:last-child { color: var(--bad, #c0392b); }
.im-batch .im-ps-value { color: var(--ink-3); font-variant-numeric: tabular-nums; }
`;
