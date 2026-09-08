// Rutba Image.
//
// Non-destructive by construction: an edit is a list of operations, the source
// bitmap is never modified, and undo is dropping the last operation. The canvas
// is redrawn from the source and the plan every time, which sounds wasteful and
// is not — one resample from the original is both faster and sharper than a
// chain of resamples through intermediate copies.
//
// Export goes through the same plan at full resolution, so what is saved is
// what was previewed, at the size of the original rather than the size of the
// window.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Field, useToast, useCommands, formatBytes, basename } from '@rutba/office-ui';
import { OPS, DEFAULT_ADJUSTMENTS, toCanvasPlan, resultSize, isIdentity, constrainCrop } from '@rutba/imaging/pipeline';
import { AppFrame, useAppMenu, pickOpen, useFileDrop, openInApp } from '../shell.js';

const fileUrl = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

const ADJUSTMENTS = [
  { key: 'brightness', label: 'Brightness', min: 0, max: 200, unit: '%' },
  { key: 'contrast', label: 'Contrast', min: 0, max: 200, unit: '%' },
  { key: 'saturation', label: 'Saturation', min: 0, max: 200, unit: '%' },
  { key: 'hue', label: 'Hue', min: -180, max: 180, unit: '°' },
  { key: 'blur', label: 'Blur', min: 0, max: 20, unit: 'px' },
  { key: 'sepia', label: 'Sepia', min: 0, max: 100, unit: '%' },
  { key: 'grayscale', label: 'Greyscale', min: 0, max: 100, unit: '%' },
];

const RATIOS = [
  { label: 'Free', value: null },
  { label: 'Square', value: 1 },
  { label: '4:3', value: 4 / 3 },
  { label: '3:2', value: 1.5 },
  { label: '16:9', value: 16 / 9 },
  { label: '9:16', value: 9 / 16 },
];

/** Draw a plan onto a canvas at a given scale. Used by preview and by export. */
function drawPlan(canvas, image, plan, scale = 1) {
  const width = Math.max(1, Math.round(plan.target.width * scale));
  const height = Math.max(1, Math.round(plan.target.height * scale));
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, width, height);
  ctx.save();
  ctx.filter = plan.filter || 'none';
  ctx.imageSmoothingQuality = 'high';
  ctx.translate(width / 2, height / 2);
  ctx.rotate((plan.rotate * Math.PI) / 180);
  ctx.scale(plan.flipX ? -1 : 1, plan.flipY ? -1 : 1);
  // After the rotation the destination box is measured in unrotated space.
  const turned = plan.rotate === 90 || plan.rotate === 270;
  const dw = turned ? height : width;
  const dh = turned ? width : height;
  ctx.drawImage(image, plan.sourceRect.x, plan.sourceRect.y, plan.sourceRect.width, plan.sourceRect.height, -dw / 2, -dh / 2, dw, dh);
  ctx.restore();

  for (const overlay of plan.overlays || []) {
    ctx.save();
    ctx.strokeStyle = overlay.colour || '#e08b2b';
    ctx.fillStyle = overlay.colour || '#e08b2b';
    ctx.lineWidth = (overlay.width || 3) * scale;
    const x = overlay.x * scale;
    const y = overlay.y * scale;
    const w = overlay.width2 * scale;
    const h = overlay.height2 * scale;
    if (overlay.shape === 'rect') ctx.strokeRect(x, y, w, h);
    else if (overlay.shape === 'ellipse') {
      ctx.beginPath();
      ctx.ellipse(x + w / 2, y + h / 2, Math.abs(w / 2), Math.abs(h / 2), 0, 0, Math.PI * 2);
      ctx.stroke();
    } else if (overlay.shape === 'text') {
      ctx.font = `${(overlay.size || 24) * scale}px "Segoe UI", system-ui, sans-serif`;
      ctx.fillText(overlay.text || '', x, y);
    }
    ctx.restore();
  }
  return canvas;
}

export default function ImageTool({ app, shell, boot }) {
  const toast = useToast();
  const [path, setPath] = useState(boot.file || null);
  const [image, setImage] = useState(null);
  const [stat, setStat] = useState(null);
  const [ops, setOps] = useState([]);
  const [adjust, setAdjust] = useState({ ...DEFAULT_ADJUSTMENTS });
  const [crop, setCrop] = useState(null);
  const [ratio, setRatio] = useState(null);
  const [tab, setTab] = useState('home');
  const [busy, setBusy] = useState(false);
  // The file that would not decode, and why. A toast says it once and fades,
  // and the window then shows "No picture open" — which is what it says when
  // you have not opened anything, so a person who just double-clicked a file
  // is told nothing happened rather than what happened.
  const [failed, setFailed] = useState(null);
  const canvasRef = useRef(null);
  const stageRef = useRef(null);
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'image', onOpen: () => openFileRef.current?.() });

  const plan = useMemo(
    () => (image ? toCanvasPlan({ width: image.naturalWidth, height: image.naturalHeight }, [...ops, { op: OPS.adjust, ...adjust }]) : null),
    [image, ops, adjust]
  );

  const load = useCallback(
    async (target) => {
      setBusy(true);
      setFailed(null);
      try {
        const info = await shell.fs.stat({ path: target });
        const img = new Image();
        img.decoding = 'sync';
        await new Promise((resolve, reject) => {
          img.onload = resolve;
          img.onerror = () => reject(new Error('This file could not be decoded as an image.'));
          img.src = fileUrl(target);
        });
        setImage(img);
        setStat(info);
        setPath(target);
        setOps([]);
        setAdjust({ ...DEFAULT_ADJUSTMENTS });
        setCrop(null);
        shell.app.addRecent({ path: target, app: 'image' }).catch(() => {});
      } catch (err) {
        setFailed(`${basename(target)}: ${err.message}`);
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [shell, toast]
  );

  useEffect(() => {
    if (boot.file) load(boot.file);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Preview: fit the plan into the stage, then draw at that scale.
  useEffect(() => {
    const canvas = canvasRef.current;
    const stage = stageRef.current;
    if (!canvas || !image || !plan || !stage) return;
    const available = { width: stage.clientWidth - 40, height: stage.clientHeight - 40 };
    const scale = Math.min(1, available.width / plan.target.width, available.height / plan.target.height);
    drawPlan(canvas, image, plan, Math.max(0.02, scale));
  }, [image, plan]);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'image');
    if (file) load(file);
  }, [shell, load]);
  openFileRef.current = openFile;

  useFileDrop(useCallback((files) => files[0] && load(files[0]), [load]));

  const applyCrop = useCallback(() => {
    if (!crop || !plan) return;
    setOps((list) => [...list, { op: OPS.crop, ...crop }]);
    setCrop(null);
  }, [crop, plan]);

  const exportAs = useCallback(
    async (format) => {
      if (!image || !plan) return;
      const target = await shell.dialog.save({
        title: 'Export picture',
        defaultPath: (path || 'picture').replace(/\.[^.]+$/, '') + (format === 'jpeg' ? '.jpg' : `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format === 'jpeg' ? 'jpg' : format] }],
      });
      if (!target) return;
      setBusy(true);
      try {
        // Full resolution, from the original pixels, through the same plan.
        const full = document.createElement('canvas');
        drawPlan(full, image, plan, 1);
        const blob = await new Promise((resolve) => full.toBlob(resolve, `image/${format}`, format === 'jpeg' ? 0.92 : undefined));
        const bytes = new Uint8Array(await blob.arrayBuffer());
        await shell.fs.write({ path: target, bytes });
        toast(`Saved ${basename(target)} — ${formatBytes(bytes.length)}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [image, plan, path, shell, toast]
  );

  const commands = useMemo(
    () => ({
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: 'Export as PNG…', icon: 'save', key: 'Mod+S', run: () => exportAs('png') },
      'edit.undo': { label: 'Undo', icon: 'undo', key: 'Mod+Z', run: () => setOps((l) => l.slice(0, -1)) },
      'edit.reset': { label: 'Reset all', icon: 'refresh', run: () => { setOps([]); setAdjust({ ...DEFAULT_ADJUSTMENTS }); setCrop(null); } },
      'image.rotateRight': { label: 'Rotate right', icon: 'rotate', key: 'Mod+R', run: () => setOps((l) => [...l, { op: OPS.rotate, degrees: 90 }]) },
      'image.rotateLeft': { label: 'Rotate left', icon: 'rotate', run: () => setOps((l) => [...l, { op: OPS.rotate, degrees: 270 }]) },
      'image.flipX': { label: 'Flip horizontally', icon: 'flip', run: () => setOps((l) => [...l, { op: OPS.flip, axis: 'x' }]) },
      'image.flipY': { label: 'Flip vertically', icon: 'flip', run: () => setOps((l) => [...l, { op: OPS.flip, axis: 'y' }]) },
    }),
    [openFile, exportAs]
  );

  useCommands(commands, [image, plan]);

  const size = image ? resultSize({ width: image.naturalWidth, height: image.naturalHeight }, ops) : null;
  const changed = ops.length > 0 || !isIdentity([{ op: OPS.adjust, ...adjust }]);

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={path ? basename(path) : 'Image'}
      dirty={changed}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[{ id: 'home', label: 'Home' }, { id: 'adjust', label: 'Adjust' }, { id: 'export', label: 'Export' }]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="undo" title="Undo" disabled={!ops.length} onClick={() => commands['edit.undo'].run()} />
              <Button icon="refresh" title="Reset" disabled={!changed} onClick={() => commands['edit.reset'].run()} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="File">
                <Button tall icon="open" label="Open" onClick={openFile} />
                <Button tall icon="save" label="Export" disabled={!image} onClick={() => exportAs('png')} />
              </Group>
              <Group label="Rotate">
                <Button tall icon="rotate" label="Right" disabled={!image} onClick={() => commands['image.rotateRight'].run()} />
                <Button tall icon="rotate" label="Left" disabled={!image} onClick={() => commands['image.rotateLeft'].run()} />
                <Button tall icon="flip" label="Flip" disabled={!image} onClick={() => commands['image.flipX'].run()} />
              </Group>
              <Group label="Crop">
                {RATIOS.map((r) => (
                  <Button
                    key={r.label}
                    label={r.label}
                    pressed={ratio === r.value}
                    disabled={!image}
                    onClick={() => {
                      setRatio(r.value);
                      const bounds = { width: plan.target.width, height: plan.target.height };
                      const base = { x: bounds.width * 0.1, y: bounds.height * 0.1, width: bounds.width * 0.8, height: bounds.height * 0.8 };
                      setCrop(constrainCrop(base, bounds, r.value));
                    }}
                  />
                ))}
                <Button icon="crop" label="Apply" primary disabled={!crop} onClick={applyCrop} />
              </Group>
            </>
          ) : tab === 'adjust' ? (
            <Group label="Adjustments">
              <div className="im-sliders">
                {ADJUSTMENTS.map((a) => (
                  <label key={a.key} className="im-slider">
                    <span>{a.label}</span>
                    <input
                      type="range"
                      min={a.min}
                      max={a.max}
                      value={adjust[a.key]}
                      onChange={(e) => setAdjust((s) => ({ ...s, [a.key]: Number(e.target.value) }))}
                      disabled={!image}
                    />
                    <span className="im-value">{adjust[a.key]}{a.unit}</span>
                  </label>
                ))}
              </div>
            </Group>
          ) : (
            <Group label="Export as">
              <Button tall icon="export" label="PNG" disabled={!image} onClick={() => exportAs('png')} />
              <Button tall icon="export" label="JPEG" disabled={!image} onClick={() => exportAs('jpeg')} />
              <Button tall icon="export" label="WebP" disabled={!image} onClick={() => exportAs('webp')} />
            </Group>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{path || 'No picture open'}</span>
          <Spacer />
          {image ? <Chip>{image.naturalWidth} × {image.naturalHeight} source</Chip> : null}
          {size ? <Chip>{size.width} × {size.height} result</Chip> : null}
          {stat ? <Chip>{formatBytes(stat.size)}</Chip> : null}
          {ops.length ? <Chip>{ops.length} edits</Chip> : null}
        </>
      }
    >
      {!image && failed && !busy ? (
        <Empty icon="image" title="This file could not be opened">
          {failed}
          <Button primary icon="open" label="Open another" onClick={openFile} style={{ marginTop: 12 }} />
        </Empty>
      ) : !image ? (
        <Empty icon="image" title={busy ? 'Opening…' : 'No picture open'}>
          {busy ? null : 'Open a picture, or drop one onto this window. Nothing is changed until you export.'}
          {busy ? <Spinner /> : <Button primary icon="open" label="Open a picture" onClick={openFile} style={{ marginTop: 12 }} />}
        </Empty>
      ) : (
        <>
          <Content>
            <div className="im-stage" ref={stageRef}>
              <div className="im-canvas-wrap">
                <canvas ref={canvasRef} className="im-canvas" />
                {crop && plan ? <CropBox crop={crop} plan={plan} onChange={(c) => setCrop(constrainCrop(c, plan.target, ratio))} /> : null}
              </div>
            </div>
          </Content>

          <Panel right width={216} title="Adjustments" resizable>
            <div className="im-panel">
              {ADJUSTMENTS.map((a) => (
                <Field key={a.key} label={`${a.label} — ${adjust[a.key]}${a.unit}`}>
                  <input
                    type="range"
                    min={a.min}
                    max={a.max}
                    value={adjust[a.key]}
                    onChange={(e) => setAdjust((s) => ({ ...s, [a.key]: Number(e.target.value) }))}
                  />
                </Field>
              ))}
              <Button label="Reset adjustments" onClick={() => setAdjust({ ...DEFAULT_ADJUSTMENTS })} />
              <div className="im-history">
                <div className="rw-panel-head" style={{ padding: '6px 0' }}>Edits</div>
                {ops.length ? (
                  ops.map((op, i) => (
                    <div key={i} className="im-op">
                      <Icon name={op.op === OPS.crop ? 'crop' : op.op === OPS.rotate ? 'rotate' : 'flip'} size={13} />
                      <span>
                        {op.op === OPS.crop ? `Crop ${Math.round(op.width)}×${Math.round(op.height)}` : op.op === OPS.rotate ? `Rotate ${op.degrees}°` : `Flip ${op.axis}`}
                      </span>
                    </div>
                  ))
                ) : (
                  <div className="rw-hint">The original is untouched.</div>
                )}
              </div>
            </div>
          </Panel>
        </>
      )}
      <style>{CSS}</style>
    </AppFrame>
  );
}

function CropBox({ crop, plan, onChange }) {
  const start = useRef(null);
  const scaleRef = useRef(1);

  const begin = (e, mode) => {
    e.preventDefault();
    e.stopPropagation();
    const canvas = e.currentTarget.closest('.im-canvas-wrap')?.querySelector('canvas');
    scaleRef.current = canvas ? canvas.clientWidth / plan.target.width : 1;
    start.current = { x: e.clientX, y: e.clientY, crop: { ...crop }, mode };
    const move = (ev) => {
      const dx = (ev.clientX - start.current.x) / scaleRef.current;
      const dy = (ev.clientY - start.current.y) / scaleRef.current;
      const c = start.current.crop;
      onChange(
        start.current.mode === 'move'
          ? { ...c, x: c.x + dx, y: c.y + dy }
          : { ...c, width: c.width + dx, height: c.height + dy }
      );
    };
    const stop = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', stop);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', stop);
  };

  const pct = (n, total) => `${(n / total) * 100}%`;
  return (
    <div
      className="im-crop"
      style={{
        left: pct(crop.x, plan.target.width),
        top: pct(crop.y, plan.target.height),
        width: pct(crop.width, plan.target.width),
        height: pct(crop.height, plan.target.height),
      }}
      onMouseDown={(e) => begin(e, 'move')}
    >
      <span className="im-crop-size">{Math.round(crop.width)} × {Math.round(crop.height)}</span>
      <span className="im-crop-handle" onMouseDown={(e) => begin(e, 'resize')} />
    </div>
  );
}

const CSS = `
.im-stage { flex: 1; min-height: 0; display: grid; place-items: center; background: var(--sunken); padding: 20px; overflow: auto; }
.im-canvas-wrap { position: relative; line-height: 0; }
.im-canvas { display: block; box-shadow: var(--shadow-2); background: #fff; max-width: 100%; }
.im-crop { position: absolute; border: 1px solid #fff; box-shadow: 0 0 0 9999px rgba(0,0,0,0.42); cursor: move; }
.im-crop-size {
  position: absolute; top: -22px; left: 0; font-size: 11px; color: #fff; background: rgba(0,0,0,0.6);
  padding: 1px 6px; border-radius: 3px; line-height: 1.6;
}
.im-crop-handle {
  position: absolute; right: -6px; bottom: -6px; width: 12px; height: 12px;
  background: #fff; border: 1px solid var(--accent); border-radius: 2px; cursor: nwse-resize;
}
.im-panel { padding: 6px 12px 16px; display: flex; flex-direction: column; gap: 10px; }
.im-panel input[type=range] { width: 100%; accent-color: var(--accent); }
.im-sliders { display: grid; grid-template-columns: repeat(2, auto); gap: 2px 14px; }
.im-slider { display: flex; align-items: center; gap: 7px; font-size: 11.5px; }
.im-slider span:first-child { width: 68px; color: var(--ink-2); }
.im-slider input { width: 90px; accent-color: var(--accent); }
.im-value { width: 40px; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.im-history { border-top: 1px solid var(--line-soft); margin-top: 4px; }
.im-op { display: flex; align-items: center; gap: 7px; padding: 3px 0; font-size: 12px; color: var(--ink-2); }
`;
