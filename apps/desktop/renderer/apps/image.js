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
// window. Tone and detail (levels, curves, white balance, sharpen and the
// rest) are worked on the pixels by @rutba/imaging's pixels.js, in the image
// worker, so a slider never makes the window stutter.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Rows, Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Select, Dialog, Input, useToast, useCommands, formatBytes, basename, t, tn } from '@rutba/office-ui';
import { OPS, DEFAULT_ADJUSTMENTS, toCanvasPlan, isIdentity, constrainCrop, markToOperation, spotToOperation, spotInTarget, resizeTo } from '@rutba/imaging/pipeline';
import { DEFAULT_PIXELS, NEUTRAL_LEVELS, CURVE_PRESETS, autoEnhance, histogram, isNeutral } from '@rutba/imaging/pixels';
import { FORMATS, formatOf, withExifFrom, underSize } from '@rutba/imaging/batch';
import { AppFrame, useAppMenu, pickOpen, useFileDrop } from '../shell.js';
import { drawBase, drawMarks } from './image/draw.js';
import { pixelsOff } from './image/work.js';
import { DrawLayer, markIsSomething } from './image/marks.js';
import { Histogram, CurveEditor } from './image/tone.js';

const fileUrl = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

// The sliders, by what they change: the canvas filter (`css`) or the pixels.
const SLIDERS = {
  brightness: { label: t('Brightness'), kind: 'css', min: 0, max: 200, unit: '%' },
  contrast: { label: t('Contrast'), kind: 'css', min: 0, max: 200, unit: '%' },
  exposure: { label: t('Exposure'), kind: 'px', min: -3, max: 3, step: 0.1, unit: '' },
  highlights: { label: t('Highlights'), kind: 'px', min: -100, max: 100 },
  shadows: { label: t('Shadows'), kind: 'px', min: -100, max: 100 },
  saturation: { label: t('Saturation'), kind: 'css', min: 0, max: 200, unit: '%' },
  hue: { label: t('Hue'), kind: 'css', min: -180, max: 180, unit: '°' },
  temperature: { label: t('Temperature'), kind: 'px', min: -100, max: 100 },
  tint: { label: t('Tint'), kind: 'px', min: -100, max: 100 },
  sharpen: { label: t('Sharpen'), kind: 'px', min: 0, max: 100 },
  denoise: { label: t('Noise reduction'), kind: 'px', min: 0, max: 100 },
  vignette: { label: t('Vignette'), kind: 'px', min: -100, max: 100 },
  blur: { label: t('Blur'), kind: 'css', min: 0, max: 20, unit: 'px' },
  sepia: { label: t('Sepia'), kind: 'css', min: 0, max: 100, unit: '%' },
  grayscale: { label: t('Greyscale'), kind: 'css', min: 0, max: 100, unit: '%' },
};

const SECTIONS = [
  { id: 'light', label: t('Light'), sliders: ['brightness', 'contrast', 'exposure', 'highlights', 'shadows'] },
  { id: 'colour', label: t('Colour'), sliders: ['saturation', 'hue', 'temperature', 'tint'] },
  { id: 'levels', label: t('Levels') },
  { id: 'curve', label: t('Curves') },
  { id: 'detail', label: t('Detail'), sliders: ['sharpen', 'denoise', 'vignette', 'blur'] },
  { id: 'effects', label: t('Effects'), sliders: ['grayscale', 'sepia'] },
];

const RATIOS = [
  { label: t('Free'), value: null },
  { label: t('Square'), value: 1 },
  { label: '4:3', value: 4 / 3 },
  { label: '3:2', value: 1.5 },
  { label: '16:9', value: 16 / 9 },
  { label: '9:16', value: 9 / 16 },
];

const TOOL_INFO = {
  pen: { label: t('Pen'), icon: 'pen', tip: t('Pen: draw freehand') },
  rect: { label: t('Rectangle'), icon: 'mkRect', tip: t('Rectangle: drag a box; Shift keeps it square') },
  ellipse: { label: t('Ellipse'), icon: 'mkEllipse', tip: t('Ellipse: drag an oval; Shift keeps it round') },
  arrow: { label: t('Arrow'), icon: 'mkArrow', tip: t('Arrow: drag from the tail to the point') },
  text: { label: t('Text'), icon: 'mkText', tip: t('Text: click where it starts, then type') },
};

const SWATCHES = ['#e0393e', '#f08c1a', '#f5d90a', '#2e9e4f', '#2f7de1', '#7b49d6', '#ffffff', '#111111'];
const WIDTHS = [1, 2, 3, 4, 6, 8, 12, 16, 24];
const TEXT_SIZES = [12, 16, 20, 24, 32, 40, 48, 64, 80, 96];
const SPOT_SIZES = [{ id: 's', px: 8, label: t('Small') }, { id: 'm', px: 13, label: t('Medium') }, { id: 'l', px: 20, label: t('Large') }];
const LONG_SIDES = [0, 4096, 3840, 2560, 1920, 1600, 1280, 1024, 800, 640, 480, 320];
const MAX_SIZES = [0, 100, 200, 500, 1024, 2048, 5120];

const curvePoints = (curve, channel) => (curve && curve[channel]) || CURVE_PRESETS.linear;

export default function ImageTool({ app, shell, boot }) {
  const toast = useToast();
  const [path, setPath] = useState(boot.file || null);
  const [pastedName, setPastedName] = useState(null);
  const [image, setImage] = useState(null);
  const [stat, setStat] = useState(null);
  const [ops, setOps] = useState([]);
  const [adjust, setAdjust] = useState({ ...DEFAULT_ADJUSTMENTS });
  const [pixels, setPixels] = useState({ ...DEFAULT_PIXELS });
  const [crop, setCrop] = useState(null);
  const [ratio, setRatio] = useState(null);
  const [tab, setTab] = useState('home');
  const [busy, setBusy] = useState(false);
  const [tool, setTool] = useState(null);
  const [draft, setDraft] = useState(null);
  const [markStyle, setMarkStyle] = useState({ colour: SWATCHES[0], width: 4, size: 32 });
  const [spotSize, setSpotSize] = useState('m');
  const [open, setOpen] = useState(() => new Set(['light', 'colour']));
  const [channel, setChannel] = useState('rgb');
  const [hist, setHist] = useState(null);
  const [resizing, setResizing] = useState(false);
  const [stage, setStage] = useState({ width: 0, height: 0 });
  const [scale, setScale] = useState(1);
  const [exportOpts, setExportOpts] = useState({ format: null, quality: 90, longSide: 0, maxKB: 0, strip: true });
  // The file that would not decode, and why. A toast says it once and fades,
  // and the window then shows "No picture open" — which is what it says when
  // you have not opened anything, so a person who just double-clicked a file
  // is told nothing happened rather than what happened.
  const [failed, setFailed] = useState(null);
  const canvasRef = useRef(null);
  const stageRef = useRef(null);
  const baseRef = useRef(null);
  const drawn = useRef(0);
  const openFileRef = useRef(null);
  const sectionRefs = useRef({});
  const appMenu = useAppMenu({ shell, appKey: 'image', onOpen: () => openFileRef.current?.() });

  const source = image ? { width: image.naturalWidth, height: image.naturalHeight } : null;
  // The geometry and every mark committed; the picture as it is.
  const plan = useMemo(
    () => (source ? toCanvasPlan(source, [...ops, { op: OPS.adjust, ...adjust }, { op: OPS.pixels, ...pixels }]) : null),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [image, ops, adjust, pixels]
  );
  // The mark being drawn, as the operation it will become, so it is seen on
  // the picture as it is made. A text being typed shows in its own box.
  const draftOp = useMemo(() => (plan && draft && draft.shape !== 'text' && markIsSomething(draft) ? markToOperation(plan, draft) : null), [plan, draft]);
  const straighten = useMemo(() => {
    for (let i = ops.length - 1; i >= 0; i--) if (ops[i].op === OPS.straighten) return ops[i].degrees;
    return 0;
  }, [ops]);

  const reset = useCallback(() => {
    setOps([]);
    setAdjust({ ...DEFAULT_ADJUSTMENTS });
    setPixels({ ...DEFAULT_PIXELS });
    setCrop(null);
    setDraft(null);
  }, []);

  const decode = useCallback(async (url) => {
    const img = new Image();
    img.decoding = 'sync';
    // The file host answers for any origin, so the canvas it is drawn on can
    // still be read for its pixels and exported.
    img.crossOrigin = 'anonymous';
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = () => reject(new Error(t('This file could not be decoded as an image.')));
      img.src = url;
    });
    return img;
  }, []);

  const load = useCallback(
    async (target) => {
      setBusy(true);
      setFailed(null);
      try {
        const info = await shell.fs.stat({ path: target });
        const img = await decode(fileUrl(target));
        setImage(img);
        setStat(info);
        setPath(target);
        setPastedName(null);
        reset();
        shell.app.addRecent({ path: target, app: 'image' }).catch(() => {});
      } catch (err) {
        setFailed(`${basename(target)}: ${err.message}`);
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [shell, toast, decode, reset]
  );

  /** A picture from the clipboard, opened as a new picture with no file yet. */
  const loadPasted = useCallback(
    async (picture) => {
      setBusy(true);
      setFailed(null);
      try {
        const url = URL.createObjectURL(new Blob([picture.bytes], { type: 'image/png' }));
        const img = await decode(url);
        setImage(img);
        setStat({ size: picture.bytes.length });
        setPath(null);
        setPastedName(t('Pasted picture'));
        reset();
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [decode, reset, toast]
  );

  const paste = useCallback(async () => {
    const picture = await shell.clipboard.readImage().catch(() => null);
    if (!picture) {
      toast(t('There is no picture on the clipboard to paste.'));
      return;
    }
    // Into this window when it has nothing open; a window of its own otherwise.
    if (!image) loadPasted(picture);
    else shell.win.create({ app: 'image', query: { paste: '1' } });
  }, [shell, toast, image, loadPasted]);

  useEffect(() => {
    if (boot.file) load(boot.file);
    else if (new URLSearchParams(window.location.search).get('paste')) {
      shell.clipboard.readImage().then((p) => p && loadPasted(p)).catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // The room the picture has, followed as the window and the panel change.
  useEffect(() => {
    const el = stageRef.current;
    if (!el) return undefined;
    const measure = () => setStage((s) => (s.width === el.clientWidth && s.height === el.clientHeight ? s : { width: el.clientWidth, height: el.clientHeight }));
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [image]);

  // Preview: the picture at the stage's scale, tone and detail in the
  // worker, then the marks. The canvas keeps the last frame until the next is
  // ready, so a slider dragged quickly never shows a blank.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !image || !plan || !stage.width) return;
    const fitScale = Math.max(0.02, Math.min(1, (stage.width - 48) / plan.target.width, (stage.height - 48) / plan.target.height));
    const dpr = window.devicePixelRatio || 1;
    const drawScale = fitScale * dpr;
    setScale(fitScale);
    const turn = ++drawn.current;
    const base = baseRef.current || (baseRef.current = document.createElement('canvas'));
    const ctx = drawBase(base, image, plan, drawScale, { readBack: true });
    let data = null;
    try {
      data = ctx.getImageData(0, 0, base.width, base.height);
    } catch {
      data = null;
    }
    if (data) setHist(histogram(data));
    const spots = plan.spots.map((s) => {
      const at = spotInTarget(plan, s);
      return { x: at.x * drawScale, y: at.y * drawScale, radius: at.radius * drawScale };
    });
    const needsPixels = data && (!isNeutral(plan.pixels) || spots.length);
    (async () => {
      const out = needsPixels ? await pixelsOff(data, plan.pixels, spots) : null;
      if (turn !== drawn.current) return;
      canvas.width = base.width;
      canvas.height = base.height;
      canvas.style.width = `${Math.round(plan.target.width * fitScale)}px`;
      canvas.style.height = `${Math.round(plan.target.height * fitScale)}px`;
      const c = canvas.getContext('2d');
      if (out) c.putImageData(out, 0, 0);
      else c.drawImage(base, 0, 0);
      drawMarks(c, plan, draftOp ? [...plan.overlays, draftOp] : plan.overlays, drawScale);
      canvas.dataset.drawn = String(turn);
    })();
  }, [image, plan, draftOp, stage]);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'image');
    if (file) load(file);
  }, [shell, load]);
  openFileRef.current = openFile;

  useFileDrop(useCallback((files) => files[0] && load(files[0]), [load]));

  const addOp = useCallback((op) => setOps((list) => [...list, op]), []);

  /** The mark being drawn, put into the edit list. */
  const commitDraft = useCallback(() => {
    setDraft((d) => {
      if (d && plan && markIsSomething(d)) setOps((list) => [...list, markToOperation(plan, d)]);
      return null;
    });
  }, [plan]);

  const chooseTool = useCallback((next) => {
    commitDraft();
    setCrop(null);
    setTool((cur) => (cur === next ? null : next));
  }, [commitDraft]);

  // A tab with no tools on it puts the tool down, and the mark drawn is kept.
  useEffect(() => {
    if (tab !== 'draw' && tool && tool !== 'redeye') { commitDraft(); setTool(null); }
    if (tab !== 'filters' && tool === 'redeye') setTool(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab]);

  const style = useCallback((patch) => {
    setMarkStyle((s) => ({ ...s, ...patch }));
    setDraft((d) => (d ? { ...d, ...patch } : d));
  }, []);

  const applyCrop = useCallback(() => {
    if (!crop || !plan) return;
    addOp({ op: OPS.crop, ...crop });
    setCrop(null);
  }, [crop, plan, addOp]);

  const setStraighten = useCallback((degrees) => {
    const value = Math.max(-45, Math.min(45, Math.round(Number(degrees) * 10) / 10)) || 0;
    // Dragging the slider changes the one straighten at the end of the list
    // rather than adding one for every step.
    setOps((list) => (list.length && list[list.length - 1].op === OPS.straighten
      ? (value ? [...list.slice(0, -1), { op: OPS.straighten, degrees: value }] : list.slice(0, -1))
      : value ? [...list, { op: OPS.straighten, degrees: value }] : list));
  }, []);

  const setSlider = useCallback((key, value) => {
    const spec = SLIDERS[key];
    if (spec.kind === 'css') setAdjust((s) => ({ ...s, [key]: value }));
    else setPixels((s) => ({ ...s, [key]: value }));
  }, []);
  const sliderValue = (key) => (SLIDERS[key].kind === 'css' ? adjust[key] : pixels[key]);
  const sliderDefault = (key) => (SLIDERS[key].kind === 'css' ? DEFAULT_ADJUSTMENTS[key] : DEFAULT_PIXELS[key]);

  const levels = { ...NEUTRAL_LEVELS, ...(pixels.levels || {}) };
  const setLevels = (patch) => setPixels((s) => ({ ...s, levels: { ...NEUTRAL_LEVELS, ...(s.levels || {}), ...patch } }));
  const setCurve = (points) => setPixels((s) => ({ ...s, curve: { ...(s.curve || {}), [channel]: points } }));

  const autoFix = useCallback(() => {
    // Worked out from the picture as it is shown before tone, then put on
    // the sliders, where each can be taken further or back.
    const base = baseRef.current;
    if (!base) return;
    const data = base.getContext('2d').getImageData(0, 0, base.width, base.height);
    const found = autoEnhance(data);
    setPixels((s) => ({ ...s, levels: found.levels, temperature: found.temperature, tint: found.tint }));
    toast(t('Enhanced: the tones stretched to the full range and the colour cast eased.'), { tone: 'good' });
  }, [toast]);

  const openSection = useCallback((id) => {
    setOpen((s) => new Set([...s, id]));
    requestAnimationFrame(() => sectionRefs.current[id]?.scrollIntoView({ block: 'nearest', behavior: 'smooth' }));
  }, []);

  /** The result at full size (or `scale` of it): picture, tone and detail, marks. */
  const renderFull = useCallback(async (s = 1) => {
    const full = document.createElement('canvas');
    const ctx = drawBase(full, image, plan, s);
    if (!isNeutral(plan.pixels) || plan.spots.length) {
      const data = ctx.getImageData(0, 0, full.width, full.height);
      const spots = plan.spots.map((sp) => {
        const at = spotInTarget(plan, sp);
        return { x: at.x * s, y: at.y * s, radius: at.radius * s };
      });
      ctx.putImageData(await pixelsOff(data, plan.pixels, spots), 0, 0);
    }
    drawMarks(ctx, plan, plan.overlays, s);
    return full;
  }, [image, plan]);

  const sourceFormat = formatOf(path || '') || 'png';
  const format = exportOpts.format || sourceFormat;

  const exportPicture = useCallback(
    async (fmt = format) => {
      if (!image || !plan) return;
      commitDraft();
      const info = FORMATS[fmt];
      const stem = path ? basename(path).replace(/\.[^.]+$/, '') : pastedName || t('Picture');
      const folder = path ? path.slice(0, path.length - basename(path).length) : `${(await shell.app.paths().catch(() => ({}))).pictures || ''}/`;
      const target = await shell.dialog.save({
        title: t('Export picture'),
        defaultPath: `${folder}${stem}.${info.ext}`,
        filters: [{ name: fmt.toUpperCase(), extensions: [info.ext] }],
      });
      if (!target) return;
      setBusy(true);
      try {
        const long = Math.max(plan.target.width, plan.target.height);
        let s = exportOpts.longSide ? Math.min(1, exportOpts.longSide / long) : 1;
        const maxBytes = exportOpts.maxKB * 1024;
        const encode = (canvas, q) => new Promise((resolve) => canvas.toBlob(resolve, info.mime, info.quality ? q : undefined))
          .then(async (blob) => new Uint8Array(await blob.arrayBuffer()));
        let canvas = await renderFull(s);
        let result;
        if (info.quality) {
          result = await underSize((q) => encode(canvas, q), maxBytes, { high: exportOpts.quality / 100 });
        } else {
          // A PNG has no quality to give: it is made smaller instead, a step at a time.
          let bytes = await encode(canvas);
          for (let i = 0; maxBytes && bytes.length > maxBytes && i < 14; i++) {
            s *= 0.85;
            canvas = await renderFull(s);
            bytes = await encode(canvas);
          }
          result = { bytes, fits: !maxBytes || bytes.length <= maxBytes };
        }
        let bytes = result.bytes;
        if (!exportOpts.strip && fmt === 'jpeg' && path && sourceFormat === 'jpeg') {
          // The camera's details kept from the original, turned upright as the pixels are.
          const original = await shell.fs.read({ path });
          bytes = withExifFrom(bytes, original.bytes);
        }
        await shell.fs.write({ path: target, bytes });
        const said = t('Saved {name}: {width} × {height}, {size}', { name: basename(target), width: canvas.width, height: canvas.height, size: formatBytes(bytes.length) });
        if (result.fits === false) toast(t('{saved}. It could not be brought under {limit}.', { saved: said, limit: formatBytes(maxBytes) }));
        else toast(said, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [image, plan, path, pastedName, shell, toast, exportOpts, renderFull, commitDraft, format, sourceFormat]
  );

  const copy = useCallback(async () => {
    if (!image || !plan) return;
    commitDraft();
    try {
      const canvas = await renderFull(1);
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
      await shell.clipboard.writeImage({ bytes: new Uint8Array(await blob.arrayBuffer()) });
      toast(t('Copied the picture, {width} × {height}', { width: canvas.width, height: canvas.height }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [image, plan, renderFull, shell, toast, commitDraft]);

  const undo = useCallback(() => {
    if (draft) setDraft(null);
    else setOps((l) => l.slice(0, -1));
  }, [draft]);

  const commands = useMemo(
    () => ({
      'file.open': { label: t('Open…'), icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: t('Export…'), icon: 'save', key: 'Mod+S', run: () => exportPicture() },
      'edit.undo': { label: t('Undo'), icon: 'undo', key: 'Mod+Z', run: undo },
      'edit.reset': { label: t('Reset all'), icon: 'refresh', run: reset },
      'edit.copy': { label: t('Copy the picture'), icon: 'copy', key: 'Mod+C', when: () => Boolean(image), run: copy },
      'edit.paste': { label: t('Paste as a new picture'), icon: 'paste', key: 'Mod+V', run: paste },
      'mark.commit': { label: t('Done'), icon: 'check', key: 'Enter', when: () => Boolean(draft), run: commitDraft },
      'mark.cancel': { label: t('Discard the mark'), icon: 'close', key: 'Escape', when: () => Boolean(draft || crop || tool), run: () => { if (draft) setDraft(null); else if (crop) setCrop(null); else setTool(null); } },
      'image.rotateRight': { label: t('Rotate right'), icon: 'rotate', key: 'Mod+R', run: () => addOp({ op: OPS.rotate, degrees: 90 }) },
      'image.rotateLeft': { label: t('Rotate left'), icon: 'rotateLeft', run: () => addOp({ op: OPS.rotate, degrees: 270 }) },
      'image.flipX': { label: t('Flip horizontally'), icon: 'flip', run: () => addOp({ op: OPS.flip, axis: 'x' }) },
      'image.flipY': { label: t('Flip vertically'), icon: 'flipVertical', run: () => addOp({ op: OPS.flip, axis: 'y' }) },
    }),
    [openFile, exportPicture, undo, reset, copy, paste, commitDraft, draft, crop, tool, image, addOp]
  );

  useCommands(commands, [image, plan, draft, crop, tool]);

  const size = plan ? plan.target : null;
  const changed = ops.length > 0 || !isIdentity([{ op: OPS.adjust, ...adjust }]) || !isNeutral(pixels);
  const title = path ? basename(path) : pastedName || t('Image');

  const slider = (key, { compact = false } = {}) => {
    const spec = SLIDERS[key];
    const value = sliderValue(key);
    const shown = spec.step && spec.step < 1 ? Number(value).toFixed(1) : value;
    return (
      <label key={key} className={compact ? 'im-rs' : 'im-ps'} data-slider={key}>
        <span className="im-ps-label">{spec.label}</span>
        <input
          type="range"
          min={spec.min}
          max={spec.max}
          step={spec.step || 1}
          value={value}
          disabled={!image}
          onChange={(e) => setSlider(key, Number(e.target.value))}
          onDoubleClick={() => setSlider(key, sliderDefault(key))}
        />
        <span className="im-ps-value">{shown}{spec.unit || ''}</span>
      </label>
    );
  };

  const toggleCss = (key) => setAdjust((s) => ({ ...s, [key]: s[key] ? 0 : 100 }));

  const ribbon = (
    <Ribbon
      tabs={[
        { id: 'home', label: t('Home') },
        { id: 'adjust', label: t('Adjust') },
        { id: 'filters', label: t('Filters') },
        { id: 'draw', label: t('Draw') },
        { id: 'export', label: t('Export') },
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="undo" title={t('Undo')} disabled={!ops.length && !draft} onClick={undo} />
          <Button icon="refresh" title={t('Reset')} disabled={!changed} onClick={reset} />
        </>
      }
    >
      {tab === 'home' ? (
        <>
          <Group label={t('File')}>
            <Button tall icon="open" label={t('Open')} onClick={openFile} />
            <Button tall icon="save" label={t('Export')} disabled={!image} onClick={() => exportPicture()} />
          </Group>
          <Group label={t('Clipboard')}>
            <Button tall icon="paste" label={t('Paste')} title={t('Paste: a picture on the clipboard, as a new picture (Ctrl+V)')} onClick={paste} />
            <Button tall icon="copy" label={t('Copy')} title={t('Copy: the edited picture, to paste anywhere (Ctrl+C)')} disabled={!image} onClick={copy} />
          </Group>
          <Group label={t('Rotate and flip')}>
            <Button tall icon="rotate" label={t('Right')} disabled={!image} onClick={() => commands['image.rotateRight'].run()} />
            <Button tall icon="rotateLeft" label={t('Left')} disabled={!image} onClick={() => commands['image.rotateLeft'].run()} />
            <Rows>
              <Button icon="flip" label={t('Flip horizontal')} disabled={!image} onClick={() => commands['image.flipX'].run()} />
              <Button icon="flipVertical" label={t('Flip vertical')} disabled={!image} onClick={() => commands['image.flipY'].run()} />
            </Rows>
          </Group>
          <Group label={t('Straighten')}>
            <div className="im-straighten">
              <input
                type="range"
                min={-45}
                max={45}
                step={0.5}
                value={straighten}
                disabled={!image}
                data-role="straighten"
                aria-label={t('Straighten')}
                onChange={(e) => setStraighten(e.target.value)}
                onDoubleClick={() => setStraighten(0)}
              />
              <div className="im-straighten-foot">
                <span className="im-ps-value">{`${straighten > 0 ? '+' : ''}${straighten}°`}</span>
                <Button icon="refresh" title={t('Straight again')} disabled={!straighten} onClick={() => setStraighten(0)} />
              </div>
            </div>
          </Group>
          <Group label={t('Crop')}>
            <Rows>
              <>
                {RATIOS.slice(0, 3).map((r) => (
                  <Button key={r.label} label={r.label} pressed={Boolean(crop) && ratio === r.value} disabled={!image} onClick={() => startCrop(r.value)} />
                ))}
              </>
              <>
                {RATIOS.slice(3).map((r) => (
                  <Button key={r.label} label={r.label} pressed={Boolean(crop) && ratio === r.value} disabled={!image} onClick={() => startCrop(r.value)} />
                ))}
              </>
            </Rows>
            <Button tall icon="crop" label={t('Apply')} primary={Boolean(crop)} disabled={!crop} onClick={applyCrop} />
          </Group>
          <Group label={t('Size')}>
            <Button tall icon="mkResize" label={t('Resize')} title={t('Resize: by pixels or by percent, the shape kept')} disabled={!image} onClick={() => setResizing(true)} />
          </Group>
        </>
      ) : tab === 'adjust' ? (
        <>
          <Group label={t('Automatic')}>
            <Button tall icon="wand" label={t('Auto enhance')} title={t('Auto enhance: stretch the tones and ease a colour cast in one press')} disabled={!image} onClick={autoFix} />
          </Group>
          <Group label={t('Light')}>
            <div className="im-rs-col">{['exposure', 'highlights', 'shadows'].map((k) => slider(k, { compact: true }))}</div>
          </Group>
          <Group label={t('Colour')}>
            <div className="im-rs-col">{['temperature', 'tint', 'saturation'].map((k) => slider(k, { compact: true }))}</div>
          </Group>
          <Group label={t('Tone')}>
            <Button tall icon="mkLevels" label={t('Levels')} title={t('Levels: set the black point, the white point and the middle tones')} disabled={!image} onClick={() => openSection('levels')} />
            <Button tall icon="mkCurve" label={t('Curves')} title={t('Curves: bend the tones by dragging a line')} disabled={!image} onClick={() => openSection('curve')} />
          </Group>
          <Group label={t('Start again')}>
            <Button tall icon="refresh" label={t('Reset adjustments')} disabled={!image} onClick={() => { setAdjust({ ...DEFAULT_ADJUSTMENTS }); setPixels({ ...DEFAULT_PIXELS }); }} />
          </Group>
        </>
      ) : tab === 'filters' ? (
        <>
          <Group label={t('Effects')}>
            <Button tall icon="mkInvert" label={t('Invert')} title={t('Invert: the colours turned to their opposites')} pressed={Boolean(adjust.invert)} disabled={!image} onClick={() => toggleCss('invert')} />
            <Button tall icon="contrast" label={t('Greyscale')} pressed={Boolean(adjust.grayscale)} disabled={!image} onClick={() => toggleCss('grayscale')} />
            <Button tall icon="mkSepia" label={t('Sepia')} pressed={Boolean(adjust.sepia)} disabled={!image} onClick={() => toggleCss('sepia')} />
          </Group>
          <Group label={t('Detail')}>
            <div className="im-rs-col">{['sharpen', 'denoise', 'blur'].map((k) => slider(k, { compact: true }))}</div>
          </Group>
          <Group label={t('Edges')}>
            <div className="im-rs-col">{slider('vignette', { compact: true })}</div>
          </Group>
          <Group label={t('Red eye')}>
            <Button tall icon="eye" label={t('Red eye')} title={t('Red eye: click each red pupil to darken it')} pressed={tool === 'redeye'} disabled={!image} onClick={() => chooseTool('redeye')} />
            <Rows>
              <Select value={spotSize} onChange={(e) => setSpotSize(e.target.value)} title={t('How large a pupil the click takes')} data-role="spot-size" style={{ width: 92 }}>
                {SPOT_SIZES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}
              </Select>
            </Rows>
          </Group>
        </>
      ) : tab === 'draw' ? (
        <>
          <Group label={t('Tools')}>
            {Object.entries(TOOL_INFO).map(([id, info]) => (
              <Button key={id} tall icon={info.icon} label={info.label} title={info.tip} pressed={tool === id} disabled={!image} onClick={() => chooseTool(id)} data-tool={id} />
            ))}
          </Group>
          <Group label={t('Colour')}>
            <div className="im-swatches">
              {SWATCHES.map((c) => (
                <button
                  key={c}
                  type="button"
                  className="im-swatch"
                  style={{ background: c }}
                  aria-pressed={markStyle.colour === c}
                  aria-label={c}
                  data-swatch={c}
                  onClick={() => style({ colour: c })}
                />
              ))}
              <label className="im-swatch more" data-tip={t('More colours')}>
                <input type="color" value={markStyle.colour} onChange={(e) => style({ colour: e.target.value })} aria-label={t('More colours')} />
              </label>
            </div>
          </Group>
          <Group label={t('Size')}>
            <Rows>
              <label className="im-pick">
                <span>{t('Width')}</span>
                <Select value={markStyle.width} onChange={(e) => style({ width: Number(e.target.value) })} data-role="mark-width" style={{ width: 70 }}>
                  {WIDTHS.map((w) => <option key={w} value={w}>{t('{size} px', { size: w })}</option>)}
                </Select>
              </label>
              <label className="im-pick">
                <span>{t('Text')}</span>
                <Select value={markStyle.size} onChange={(e) => style({ size: Number(e.target.value) })} data-role="text-size" style={{ width: 70 }}>
                  {TEXT_SIZES.map((w) => <option key={w} value={w}>{t('{size} px', { size: w })}</option>)}
                </Select>
              </label>
            </Rows>
          </Group>
          <Group label={t('Mark')}>
            <Button tall icon="check" label={t('Done')} title={t('Done: keep the mark as it is (Enter)')} primary={Boolean(draft)} disabled={!draft} onClick={commitDraft} />
            <Button tall icon="close" label={t('Discard')} title={t('Discard: the mark being drawn (Esc)')} disabled={!draft} onClick={() => setDraft(null)} />
          </Group>
        </>
      ) : (
        <>
          <Group label={t('Format')}>
            {['png', 'jpeg', 'webp'].map((f) => (
              <Button key={f} tall icon="export" label={f === 'webp' ? 'WebP' : f.toUpperCase()} pressed={format === f} disabled={!image} onClick={() => setExportOpts((o) => ({ ...o, format: f }))} data-format={f} />
            ))}
          </Group>
          <Group label={t('Quality')}>
            <div className="im-rs-col">
              <label className="im-rs" data-slider="quality">
                <span className="im-ps-label">{t('Quality')}</span>
                <input type="range" min={10} max={100} value={exportOpts.quality} disabled={!image || !FORMATS[format].quality} data-role="quality" onChange={(e) => setExportOpts((o) => ({ ...o, quality: Number(e.target.value) }))} />
                <span className="im-ps-value">{FORMATS[format].quality ? `${exportOpts.quality}%` : ''}</span>
              </label>
              <div className="rw-hint im-format-hint">{FORMATS[format].quality ? t('Lower is a smaller file.') : t('PNG keeps every pixel.')}</div>
            </div>
          </Group>
          <Group label={t('Limits')}>
            <Rows>
              <label className="im-pick">
                <span>{t('Longest side')}</span>
                <Select value={exportOpts.longSide} disabled={!image} data-role="long-side" onChange={(e) => setExportOpts((o) => ({ ...o, longSide: Number(e.target.value) }))} style={{ width: 112 }}>
                  {LONG_SIDES.map((v) => <option key={v} value={v}>{v ? t('{size} px', { size: v }) : t('As it is')}</option>)}
                </Select>
              </label>
              <label className="im-pick">
                <span>{t('File size')}</span>
                <Select value={exportOpts.maxKB} disabled={!image} data-role="max-size" onChange={(e) => setExportOpts((o) => ({ ...o, maxKB: Number(e.target.value) }))} style={{ width: 112 }}>
                  {MAX_SIZES.map((v) => <option key={v} value={v}>{v ? t('At most {size}', { size: formatBytes(v * 1024) }) : t('No limit')}</option>)}
                </Select>
              </label>
            </Rows>
          </Group>
          <Group label={t('Privacy')}>
            <Button
              tall
              icon="shield"
              label={t('Strip metadata')}
              title={t('Strip metadata: no camera, date or place in the file. Off, a JPEG from a JPEG keeps them.')}
              pressed={exportOpts.strip}
              disabled={!image}
              data-role="strip"
              onClick={() => setExportOpts((o) => ({ ...o, strip: !o.strip }))}
            />
          </Group>
          <Group label={t('Save')}>
            <Button tall primary icon="save" label={t('Export…')} disabled={!image} onClick={() => exportPicture()} />
          </Group>
        </>
      )}
    </Ribbon>
  );

  function startCrop(value) {
    commitDraft();
    setTool(null);
    setRatio(value);
    const bounds = { width: plan.target.width, height: plan.target.height };
    const base = { x: bounds.width * 0.1, y: bounds.height * 0.1, width: bounds.width * 0.8, height: bounds.height * 0.8 };
    setCrop(constrainCrop(base, bounds, value));
  }

  const describe = (op) => {
    switch (op.op) {
      case OPS.crop: return [t('Crop {width}×{height}', { width: Math.round(op.width), height: Math.round(op.height) }), 'crop'];
      case OPS.rotate: return [t('Rotate {degrees}°', { degrees: op.degrees }), op.degrees === 270 ? 'rotateLeft' : 'rotate'];
      case OPS.flip: return [op.axis === 'y' ? t('Flip vertical') : t('Flip horizontal'), op.axis === 'y' ? 'flipVertical' : 'flip'];
      case OPS.resize: return [t('Resize to {width}×{height}', { width: op.width, height: op.height }), 'mkResize'];
      case OPS.straighten: return [t('Straighten {degrees}°', { degrees: op.degrees }), 'mkStraighten'];
      case OPS.redeye: return [t('Red eye'), 'eye'];
      case OPS.annotate: return [TOOL_INFO[op.shape]?.label || t('Mark'), TOOL_INFO[op.shape]?.icon || 'pen'];
      default: return [op.op, 'sliders'];
    }
  };

  const spotPx = SPOT_SIZES.find((s) => s.id === spotSize)?.px || 13;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={title}
      dirty={changed}
      menu={appMenu}
      ribbon={ribbon}
      status={
        <>
          <span>{path || pastedName || t('No picture open')}</span>
          <Spacer />
          {tool ? <Chip>{tool === 'redeye' ? t('Click each red pupil') : TOOL_INFO[tool]?.tip}</Chip> : null}
          {image ? <Chip>{t('{width} × {height} source', { width: image.naturalWidth, height: image.naturalHeight })}</Chip> : null}
          {size ? <Chip>{t('{width} × {height} result', { width: size.width, height: size.height })}</Chip> : null}
          {straighten ? <Chip>{t('Straightened {degrees}°', { degrees: straighten })}</Chip> : null}
          {stat ? <Chip>{formatBytes(stat.size)}</Chip> : null}
          {ops.length ? <Chip>{tn(ops.length, '{count} edit', '{count} edits')}</Chip> : null}
        </>
      }
    >
      {!image && failed && !busy ? (
        <Empty icon="image" title={t('This file could not be opened')}>
          {failed}
          <Button primary icon="open" label={t('Open another')} onClick={openFile} style={{ marginTop: 12 }} />
        </Empty>
      ) : !image ? (
        <Empty icon="image" title={busy ? t('Opening…') : t('No picture open')}>
          {busy ? null : t('Open a picture, paste one, or drop one onto this window. Nothing is changed until you export.')}
          {busy ? <Spinner /> : (
            <div className="im-empty-actions">
              <Button primary icon="open" label={t('Open a picture')} onClick={openFile} />
              <Button icon="paste" label={t('Paste')} onClick={paste} />
            </div>
          )}
        </Empty>
      ) : (
        <>
          <Content>
            <div className="im-stage" ref={stageRef}>
              <div className="im-canvas-wrap" data-result={size ? `${size.width}x${size.height}` : ''}>
                <canvas ref={canvasRef} className="im-canvas" />
                {crop && plan ? <CropBox crop={crop} plan={plan} onChange={(c) => setCrop(constrainCrop(c, plan.target, ratio))} /> : null}
                {tool && plan ? (
                  <DrawLayer
                    tool={tool}
                    draft={draft}
                    onDraft={setDraft}
                    onCommit={commitDraft}
                    scale={scale}
                    style={markStyle}
                    spotRadius={spotPx}
                    onSpot={(spot) => addOp(spotToOperation(plan, spot))}
                  />
                ) : null}
              </div>
            </div>
          </Content>

          <Panel right width={252} title={t('Adjust')} resizable>
            <div className="im-panel">
              {SECTIONS.map((sec) => (
                <section key={sec.id} className="im-sec" data-section={sec.id} ref={(el) => { sectionRefs.current[sec.id] = el; }}>
                  <button
                    type="button"
                    className="im-sec-head"
                    aria-expanded={open.has(sec.id)}
                    onClick={() => setOpen((s) => { const n = new Set(s); if (n.has(sec.id)) n.delete(sec.id); else n.add(sec.id); return n; })}
                  >
                    <Icon name={open.has(sec.id) ? 'chevronDown' : 'chevronRight'} size={12} />
                    <span>{sec.label}</span>
                  </button>
                  {open.has(sec.id) ? (
                    <div className="im-sec-body">
                      {sec.sliders ? sec.sliders.map((k) => slider(k)) : null}
                      {sec.id === 'effects' ? (
                        <label className="im-check">
                          <input type="checkbox" checked={Boolean(adjust.invert)} onChange={() => toggleCss('invert')} />
                          <span>{t('Invert the colours')}</span>
                        </label>
                      ) : null}
                      {sec.id === 'levels' ? (
                        <div className="im-levels">
                          <Histogram hist={hist} marks={levels} />
                          {[
                            ['inBlack', t('Black point'), 0, 254, 1],
                            ['gamma', t('Middle tones'), -2, 2, 0.01],
                            ['inWhite', t('White point'), 1, 255, 1],
                            ['outBlack', t('Darkest output'), 0, 255, 1],
                            ['outWhite', t('Lightest output'), 0, 255, 1],
                          ].map(([key, label, min, max, step]) => (
                            <label key={key} className="im-ps" data-level={key}>
                              <span className="im-ps-label">{label}</span>
                              <input
                                type="range"
                                min={min}
                                max={max}
                                step={step}
                                // Gamma on a doubling scale, so 1 sits in the middle and 0.5 and 2 are as far either side.
                                value={key === 'gamma' ? Math.log2(levels.gamma) : levels[key]}
                                onChange={(e) => setLevels({ [key]: key === 'gamma' ? Math.round(2 ** Number(e.target.value) * 100) / 100 : Number(e.target.value) })}
                                onDoubleClick={() => setLevels({ [key]: NEUTRAL_LEVELS[key] })}
                              />
                              <span className="im-ps-value">{key === 'gamma' ? Number(levels[key]).toFixed(2) : levels[key]}</span>
                            </label>
                          ))}
                          <div className="im-sec-actions">
                            <Button label={t('Auto')} onClick={autoFix} />
                            <Button label={t('Reset levels')} onClick={() => setPixels((s) => ({ ...s, levels: null }))} />
                          </div>
                        </div>
                      ) : null}
                      {sec.id === 'curve' ? (
                        <div className="im-curves">
                          <div className="im-curve-bar">
                            <Select value={channel} onChange={(e) => setChannel(e.target.value)} data-role="curve-channel">
                              <option value="rgb">{t('All colours')}</option>
                              <option value="r">{t('Red')}</option>
                              <option value="g">{t('Green')}</option>
                              <option value="b">{t('Blue')}</option>
                            </Select>
                            <Select
                              value=""
                              data-role="curve-preset"
                              onChange={(e) => e.target.value && setCurve(CURVE_PRESETS[e.target.value].map((p) => [...p]))}
                            >
                              <option value="">{t('Presets')}</option>
                              <option value="lighter">{t('Lighter')}</option>
                              <option value="darker">{t('Darker')}</option>
                              <option value="contrast">{t('More contrast')}</option>
                              <option value="softer">{t('Less contrast')}</option>
                              <option value="linear">{t('Straight line')}</option>
                            </Select>
                          </div>
                          <CurveEditor
                            points={curvePoints(pixels.curve, channel)}
                            onChange={setCurve}
                            hist={hist}
                            tone={{ rgb: 'var(--ink)', r: '#d23c3c', g: '#2e9e4f', b: '#2f6fe0' }[channel]}
                          />
                          <div className="rw-hint">{t('Click the line to add a point, drag it to bend the tones; drag a point off the box, or double-click it, to take it away.')}</div>
                        </div>
                      ) : null}
                    </div>
                  ) : null}
                </section>
              ))}
              <div className="im-history">
                <div className="rw-panel-head" style={{ padding: '6px 0' }}>{t('Edits')}</div>
                {ops.length ? (
                  ops.map((op, i) => {
                    const [label, icon] = describe(op);
                    return (
                      <div key={i} className="im-op" data-op={op.op}>
                        <Icon name={icon} size={13} />
                        <span>{label}</span>
                      </div>
                    );
                  })
                ) : (
                  <div className="rw-hint">{t('The original is untouched.')}</div>
                )}
              </div>
            </div>
          </Panel>
        </>
      )}
      {resizing && plan ? (
        <ResizeDialog
          size={plan.target}
          onClose={() => setResizing(false)}
          onApply={(next) => {
            setResizing(false);
            if (next.width !== plan.target.width || next.height !== plan.target.height) addOp({ op: OPS.resize, ...next });
          }}
        />
      ) : null}
      <style>{CSS}</style>
    </AppFrame>
  );
}

/** Resize: by pixels or percent, the ratio kept unless it is let go, as Paint's dialog asks. */
function ResizeDialog({ size, onApply, onClose }) {
  const [by, setBy] = useState('percent');
  const [keep, setKeep] = useState(true);
  const [vals, setVals] = useState({ width: 100, height: 100 });
  const [changed, setChanged] = useState('width');
  const next = resizeTo(size, { by, width: by === 'percent' && keep ? vals[changed] : vals.width, height: vals.height, keep, changed });
  const switchTo = (mode) => {
    setBy(mode);
    setVals(mode === 'percent' ? { width: 100, height: 100 } : { width: size.width, height: size.height });
  };
  // With the shape kept, the side not being typed follows the one that is.
  const shown = (which) => {
    if (!keep || which === changed) return vals[which];
    if (by === 'percent') return vals[changed];
    return which === 'width' ? next.width : next.height;
  };
  const field = (which) => (
    <Input
      type="number"
      min={1}
      value={shown(which)}
      data-role={`resize-${which}`}
      onChange={(e) => { const v = e.target.value; setChanged(which); setVals((o) => ({ ...o, [which]: v })); }}
      style={{ width: 90 }}
    />
  );
  return (
    <Dialog
      title={t('Resize')}
      width={360}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Resize')} onClick={() => onApply(next)} data-role="resize-ok" />
        </>
      }
    >
      <div className="im-resize">
        <div className="im-resize-by" role="radiogroup" aria-label={t('Resize by')}>
          <label><input type="radio" name="by" checked={by === 'percent'} onChange={() => switchTo('percent')} data-role="resize-percent" /> {t('Percent')}</label>
          <label><input type="radio" name="by" checked={by === 'pixels'} onChange={() => switchTo('pixels')} data-role="resize-pixels" /> {t('Pixels')}</label>
        </div>
        <div className="im-resize-grid">
          <span>{t('Width')}</span>
          {field('width')}
          <span>{by === 'percent' ? '%' : t('px')}</span>
          <span>{t('Height')}</span>
          {field('height')}
          <span>{by === 'percent' ? '%' : t('px')}</span>
        </div>
        <label className="im-check">
          <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} />
          <span>{t('Keep the shape (aspect ratio)')}</span>
        </label>
        <div className="rw-hint">{t('Now {width} × {height}; after, {newWidth} × {newHeight}.', { width: size.width, height: size.height, newWidth: next.width, newHeight: next.height })}</div>
      </div>
    </Dialog>
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
      <span className="im-crop-grid" />
      <span className="im-crop-handle" onMouseDown={(e) => begin(e, 'resize')} />
    </div>
  );
}

const CSS = `
.im-stage { flex: 1; min-height: 0; display: grid; place-items: center; background: var(--sunken); padding: 24px; overflow: auto; }
.im-canvas-wrap { position: relative; line-height: 0; }
.im-canvas { display: block; box-shadow: var(--shadow-2); background:
  repeating-conic-gradient(#e9e9ec 0% 25%, #ffffff 0% 50%) 50% / 16px 16px; }
.im-crop { position: absolute; border: 1px solid #fff; box-shadow: 0 0 0 9999px rgba(0,0,0,0.42); cursor: move; }
.im-crop-grid { position: absolute; inset: 0; pointer-events: none;
  background:
    linear-gradient(to right, transparent calc(33.33% - 0.5px), rgba(255,255,255,0.5) calc(33.33% - 0.5px), rgba(255,255,255,0.5) calc(33.33% + 0.5px), transparent calc(33.33% + 0.5px), transparent calc(66.66% - 0.5px), rgba(255,255,255,0.5) calc(66.66% - 0.5px), rgba(255,255,255,0.5) calc(66.66% + 0.5px), transparent calc(66.66% + 0.5px)),
    linear-gradient(to bottom, transparent calc(33.33% - 0.5px), rgba(255,255,255,0.5) calc(33.33% - 0.5px), rgba(255,255,255,0.5) calc(33.33% + 0.5px), transparent calc(33.33% + 0.5px), transparent calc(66.66% - 0.5px), rgba(255,255,255,0.5) calc(66.66% - 0.5px), rgba(255,255,255,0.5) calc(66.66% + 0.5px), transparent calc(66.66% + 0.5px)); }
.im-crop-size {
  position: absolute; top: -22px; left: 0; font-size: 11px; color: #fff; background: rgba(0,0,0,0.6);
  padding: 1px 6px; border-radius: 3px; line-height: 1.6;
}
.im-crop-handle {
  position: absolute; right: -6px; bottom: -6px; width: 12px; height: 12px;
  background: #fff; border: 1px solid var(--accent); border-radius: 2px; cursor: nwse-resize;
}
.im-draw { position: absolute; inset: 0; cursor: crosshair; line-height: normal; }
.im-draw.tool-text { cursor: text; }
.im-draw.tool-redeye { cursor: none; overflow: hidden; }
.im-ring { position: absolute; left: 0; top: 0; border-radius: 50%; pointer-events: none; opacity: 0;
  border: 1.5px solid #fff; box-shadow: 0 0 0 1px rgba(0,0,0,0.55), inset 0 0 0 1px rgba(0,0,0,0.35); }
.im-draft { position: absolute; border: 1px dashed var(--accent); border-radius: 2px; cursor: move;
  box-shadow: 0 0 0 1px rgba(255,255,255,0.7); }
.im-draft.text { border-style: solid; background: rgba(255,255,255,0.08); cursor: default; padding: 0; display: flex; }
.im-text-grip { position: absolute; left: -9px; top: -9px; width: 14px; height: 14px; border-radius: 50%;
  background: var(--accent); border: 2px solid #fff; cursor: move; box-shadow: var(--shadow-1); }
.im-text-input { border: 0; outline: 0; resize: none; background: transparent; padding: 0; margin: 0; overflow: hidden;
  min-width: 60px; caret-color: currentColor; }
.im-text-input::placeholder { color: rgba(128,128,128,0.8); }
.im-handle { position: absolute; width: 10px; height: 10px; margin: -5px 0 0 -5px; background: #fff;
  border: 1.5px solid var(--accent); border-radius: 2px; box-shadow: var(--shadow-1); }
.im-handle.h-nw, .im-handle.h-se { cursor: nwse-resize; }
.im-handle.h-ne, .im-handle.h-sw { cursor: nesw-resize; }
.im-handle.h-a, .im-handle.h-b { border-radius: 50%; cursor: grab; }
.im-empty-actions { display: flex; gap: 8px; justify-content: center; margin-top: 12px; }

/* the ribbon's sliders, three to a group */
/* Each group's sliders share one grid, so their labels and tracks line up. */
.im-rs-col { display: grid; grid-template-columns: max-content 96px 40px; align-items: center; column-gap: 8px; row-gap: 1px; align-self: center; }
.im-rs { display: contents; font-size: 11.5px; }
.im-rs > * { height: 19px; display: flex; align-items: center; justify-content: flex-start; }
.im-rs .im-ps-value { justify-content: flex-end; }
.im-rs .im-ps-label { color: var(--ink-2); white-space: nowrap; font-size: 11.5px; }
.im-format-hint { grid-column: 1 / -1; }

/* Sliders drawn the same everywhere in the window: a fine track, a round thumb. */
.im-rs input[type=range], .im-ps input[type=range], .im-straighten input[type=range] {
  -webkit-appearance: none; appearance: none; width: 100%; height: 16px; margin: 0; background: transparent; cursor: pointer; }
.im-rs input[type=range]::-webkit-slider-runnable-track, .im-ps input[type=range]::-webkit-slider-runnable-track,
.im-straighten input[type=range]::-webkit-slider-runnable-track { height: 3px; border-radius: 2px; background: var(--line); }
.im-rs input[type=range]::-webkit-slider-thumb, .im-ps input[type=range]::-webkit-slider-thumb,
.im-straighten input[type=range]::-webkit-slider-thumb { -webkit-appearance: none; width: 13px; height: 13px; margin-top: -5px;
  border-radius: 50%; background: var(--bg, #fff); border: 2px solid var(--accent); box-shadow: 0 1px 2px rgba(0,0,0,0.18); }
input[type=range]:disabled::-webkit-slider-thumb { border-color: var(--line); }
.im-rs input[type=range]:focus-visible, .im-ps input[type=range]:focus-visible, .im-straighten input[type=range]:focus-visible { outline: 2px solid var(--accent-line, var(--accent)); outline-offset: 2px; border-radius: 4px; }
.im-ps-value { color: var(--ink-3); font-variant-numeric: tabular-nums; font-size: 11.5px; text-align: end; white-space: nowrap; }
.im-format-hint { font-size: 11px; }
.im-straighten { display: flex; flex-direction: column; gap: 4px; width: 132px; align-self: center; }
.im-straighten-foot { display: flex; align-items: center; justify-content: space-between; }
.im-straighten-foot .rw-btn.icon { width: 24px; height: 22px; }
.im-pick { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: var(--ink-2); }
.im-pick > span { min-width: 62px; }
.im-swatches { display: grid; grid-template-columns: repeat(5, 20px); gap: 5px; align-self: center; }
.im-swatch { width: 20px; height: 20px; border-radius: 50%; border: 1px solid rgba(0,0,0,0.22); padding: 0; cursor: pointer;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.25); position: relative; }
.im-swatch[aria-pressed='true'] { box-shadow: 0 0 0 2px var(--bg, #fff), 0 0 0 3.5px var(--accent); }
.im-swatch.more { background: conic-gradient(#e0393e, #f5d90a, #2e9e4f, #2f7de1, #7b49d6, #e0393e); overflow: hidden; }
.im-swatch.more input { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }

/* the panel */
.im-panel { padding: 2px 10px 16px; display: flex; flex-direction: column; }
.im-sec { border-bottom: 1px solid var(--line-soft); }
.im-sec-head { display: flex; align-items: center; gap: 6px; width: 100%; padding: 8px 2px; background: none; border: 0;
  color: var(--ink); font: inherit; font-size: 12px; font-weight: 600; cursor: pointer; text-align: start; }
.im-sec-head:hover { color: var(--accent); }
.im-sec-body { display: flex; flex-direction: column; gap: 8px; padding: 0 2px 12px; }
.im-ps { display: grid; grid-template-columns: 1fr auto; grid-template-areas: 'label value' 'range range'; row-gap: 3px; font-size: 12px; }
.im-ps .im-ps-label { grid-area: label; color: var(--ink-2); }
.im-ps .im-ps-value { grid-area: value; }
.im-ps input[type=range] { grid-area: range; }
.im-check { display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--ink-2); }
.im-levels, .im-curves { display: flex; flex-direction: column; gap: 8px; }
.im-hist { width: 100%; height: 56px; display: block; color: var(--ink-3); background: var(--sunken); border-radius: 3px; }
.im-sec-actions { display: flex; gap: 6px; }
.im-curve-bar { display: flex; gap: 6px; }
.im-curve-bar .rw-select { flex: 1; min-width: 0; }
.im-curve { width: 100%; aspect-ratio: 1; display: block; touch-action: none; cursor: crosshair; }
.im-curve-bg { fill: var(--sunken); stroke: var(--line); stroke-width: 1; }
.im-curve-hist rect { fill: var(--ink-3); opacity: 0.22; }
.im-curve-grid line { stroke: var(--line-soft); stroke-width: 1; }
.im-curve-ref { stroke: var(--line); stroke-width: 1; stroke-dasharray: 3 3; }
.im-curve-line { fill: none; stroke-width: 2.2; }
.im-curve-point { fill: var(--bg, #fff); stroke: var(--accent); stroke-width: 2; cursor: grab; }
.im-history { margin-top: 8px; }
.im-op { display: flex; align-items: center; gap: 7px; padding: 3px 0; font-size: 12px; color: var(--ink-2); }

/* Resize */
.im-resize { display: flex; flex-direction: column; gap: 12px; }
.im-resize-by { display: flex; gap: 18px; font-size: 12.5px; }
.im-resize-by label { display: flex; align-items: center; gap: 6px; }
.im-resize-grid { display: grid; grid-template-columns: 64px 96px auto; align-items: center; gap: 8px 10px; font-size: 12.5px; color: var(--ink-2); }
`;
