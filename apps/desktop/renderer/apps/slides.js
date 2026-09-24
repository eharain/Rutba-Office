// Rutba Presentation.
//
// The slide you are looking at is drawn by the engine as SVG and sent here as
// one string. That sounds lazy and is the opposite: the same renderer produces
// the editor canvas, the sorter thumbnails and the printed page, so a slide
// cannot look one way while you edit it and another way when it is shown.
//
// Editing is direct: click a text box and type into it in place, drag it to
// move it. Each change is an operation on the deck in the backend, which
// rewrites one slide's XML and leaves every other part of the file alone.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Dialog, Field, ZoomSlider, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { PrintDialog, defaultPrintOptions } from '../print.js';
import { SITE } from '@rutba/office-formats/registry';
import Presenter from './slides/presenter.js';
import SlidesRibbon from './slides/ribbon.js';

export default function Slides({ app, shell, boot }) {
  // A presenter window is the same app pointed at the same open document,
  // told to draw the speaker's side of it. It is a window rather than a panel
  // so it can live on the other screen, which is the whole point.
  const presenterFor = new URLSearchParams(location.search).get('presenter');
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [index, setIndex] = useState(0);
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState('home');
  const [editing, setEditing] = useState(null);
  /** The shape clipboard: one shape, copied in this window, pasted on any slide of it. */
  const [clip, setClip] = useState(null);
  /** The Format Painter, armed with a shape's look: its fill, its outline and its first run's font. */
  const [painter, setPainter] = useState(null);
  const [present, setPresent] = useState(false);
  const [notesOpen, setNotesOpen] = useState(false);
  /** The Header & Footer dialog, open with a box pre-ticked ('date' | 'number') or as it stands. */
  const [footerOpen, setFooterOpen] = useState(null);
  /** Find and replace across the deck. */
  const [findOpen, setFindOpen] = useState(false);
  /** Insert → Link on the selected shape. */
  const [linkOpen, setLinkOpen] = useState(false);
  /** Home → Section → Rename: the section (by index) and the name it has now. */
  const [sectionRename, setSectionRename] = useState(null);
  /** A shape to select once the slide a match is on has been shown. */
  const pendingSelect = useRef(null);
  const [blank, setBlank] = useState(false);
  // Set when a presenter window is driving, so this one follows rather than leads.
  const [led, setLed] = useState(false);
  /**
   * How the deck is shown. None of it is in the file: the view mode, the
   * rulers, gridlines and guides, whether the notes strip shows, a zoom
   * level (null fits the window), and the colour/greyscale tone.
   */
  const [view, setView] = useState({ mode: 'normal', ruler: false, gridlines: false, guides: false, notes: true, zoom: null, tone: 'colour', pane: null });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  // The selected shape — one click selects, a double-click edits its words —
  // is what the Font and Paragraph groups act on.
  const [selected, setSelected] = useState(null);
  // A drag in progress on the stage: the shape, whether it is moved or
  // resized (and by which handle), and the box it has been dragged to, in
  // slide units. The engine is told once, on release.
  const [drag, setDrag] = useState(null);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [printing, setPrinting] = useState(false);
  // Reading View: the show in this window, without going full screen.
  const [reading, setReading] = useState(false);
  const stageRef = useRef(null);
  const dragRef = useRef(null);
  const actRef = useRef(null);

  /**
   * A press on a shape or a handle starts a drag: the pointer's travel,
   * divided by the stage's scale, moves the box or resizes it — a corner
   * keeps a picture's proportions, Shift keeps any shape's. The box follows
   * the pointer as a dashed outline; the engine is told once, on release.
   */
  const startDrag = (e, shape, kind, handle = null) => {
    if (e.button !== 0 || !shape.geometry) return;
    e.preventDefault();
    e.stopPropagation();
    const g0 = { ...shape.geometry };
    const s = dragRef.current?.scale ?? 1;
    const x0 = e.clientX;
    const y0 = e.clientY;
    let moved = false;
    let g = g0;
    setSelected(shape.id);
    const move = (ev) => {
      const dx = (ev.clientX - x0) / s;
      const dy = (ev.clientY - y0) / s;
      if (!moved && Math.abs(ev.clientX - x0) < 2 && Math.abs(ev.clientY - y0) < 2) return;
      moved = true;
      g = kind === 'move' ? { ...g0, x: g0.x + dx, y: g0.y + dy } : resized(g0, handle, dx, dy, ev.shiftKey || shape.kind === 'picture');
      setDrag({ id: shape.id, kind, handle, g });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (moved) actRef.current?.('dragEnd', { id: shape.id, g });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  // The slide is drawn at its own size and scaled to fit the stage, the way
  // PowerPoint's "Fit to Window" does — a 1280-px slide in a 1000-px stage
  // used to run off the right edge, logo and all. Re-measured on resize.
  const [fit, setFit] = useState(1);
  useEffect(() => {
    const el = stageRef.current;
    if (!el || !model?.size) return undefined;
    const measure = () => {
      const w = el.clientWidth - 44;
      const h = el.clientHeight - 44;
      const scale = Math.min(1, w / model.size.width, h / model.size.height);
      setFit(Number.isFinite(scale) && scale > 0 ? Math.round(scale * 1000) / 1000 : 1);
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [model?.size?.width, model?.size?.height, busy]);
  const menu = useMenu();
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'slides', onNew: () => shell.win.create({ app: 'slides' }), onOpen: () => openFileRef.current?.() });

  const load = useCallback(
    async (next = index) => {
      if (!doc) return;
      const m = await shell.doc.model({ id: doc.id, slide: next, width: 1280 });
      setModel(m);
    },
    [doc, index, shell]
  );

  // The thumbnails the model left out arrive after the first paint, a few at
  // a time: a nineteen-slide deck shows its slide before its sidebar is
  // complete rather than after. A slide whose thumbnail is missing shows its
  // title until then.
  const missingThumbs = model?.outline ? model.outline.filter((o) => o.thumbnail == null).length : 0;
  useEffect(() => {
    if (!doc || !missingThumbs) return undefined;
    let alive = true;
    (async () => {
      const missing = (model?.outline || []).filter((o) => o.thumbnail == null).map((o) => o.index);
      for (let at = 0; at < missing.length && alive; at += 4) {
        let got = {};
        try {
          got = await shell.doc.thumbnails({ id: doc.id, indexes: missing.slice(at, at + 4) });
        } catch {
          return;
        }
        if (!alive) return;
        setModel((m) => {
          if (!m?.outline) return m;
          let changed = false;
          const outline = m.outline.map((o) => {
            if (o.thumbnail == null && got[o.index]) {
              changed = true;
              return { ...o, thumbnail: got[o.index] };
            }
            return o;
          });
          return changed ? { ...m, outline } : m;
        });
      }
    })();
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc?.id, missingThumbs, shell]);


  const apply = useCallback(
    async (...ops) => {
      if (!doc) return;
      try {
        const next = await shell.doc.apply({ id: doc.id, ops, slide: index, width: 1280 });
        setDoc(next);
        setModel(next.model);
        // Returned, not swallowed: adding a slide needs to know which one it
        // got so it can select it.
        return next;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return null;
      }
    },
    [doc, index, shell, toast]
  );

  /**
   * A picture from this device onto the current slide. The service reads the
   * size out of the picture's own header and fits it to the slide; what comes
   * back is selected, so Arrange and Delete act on it at once.
   */
  const insertPicture = useCallback(async () => {
    const [file] = await shell.dialog.open({
      title: 'Insert picture',
      filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp'] }],
    });
    if (!file) return;
    const { bytes, stat } = await shell.fs.read({ path: file });
    const ext = String(stat?.ext || file.split('.').pop()).replace('.', '').toLowerCase();
    const contentType = { png: 'image/png', gif: 'image/gif', bmp: 'image/bmp' }[ext] || 'image/jpeg';
    const next = await apply({ op: 'addPicture', slide: index, name: stat?.name || file.split(/[\/]/).pop(), contentType, data: bytes });
    const added = next?.model?.slide?.shapes?.slice(-1)[0];
    if (added) setSelected(added.id);
  }, [shell, apply, index]);


  useEffect(() => {
    const template = new URLSearchParams(location.search).get('template');
    const run = async () => {
      setBusy(true);
      try {
        const recover = new URLSearchParams(location.search).get('recover');
        const opened = recover
          ? await shell.doc.recover({ file: recover })
          : boot.file
            ? await shell.doc.open({ path: boot.file, kind: 'deck', width: 1280 })
            : await shell.doc.new({ kind: 'slides', template: template && template !== 'blank' ? template : 'deck' });
        if (recover) toast('Recovered unsaved work. Save it to keep it.', { ms: 6000 });
        setDoc(opened);
        setModel(opened.model);
        if (opened.path) shell.app.addRecent({ path: opened.path, app: 'slides' }).catch(() => {});
        // What saving will actually do, which is not one answer: a .md
        // opened here saves as .md, and an .rtf cannot be saved at all
        // until it is given a new name. Both used to be promised a .pptx.
        if (opened.converted?.from) {
          const was = opened.converted.from.toUpperCase();
          toast(
            opened.converted.writesBack
              ? `Opened from ${was}. Saving writes the ${was} back.`
              : `Opened from ${was}. This build cannot write ${was} — Save as will write a .pptx.`,
            { ms: 5200 }
          );
        }
      } catch (err) {
        setError(err.message);
      } finally {
        setBusy(false);
      }
    };
    run();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    load(index);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [index, doc]);

  // When a presenter window is driving, this one follows: the position lives
  // in the main process precisely so the two cannot disagree about it.
  useEffect(() => {
    const off = shell.on('present:state', (s) => {
      setLed(true);
      setBlank(Boolean(s.blank));
      if (typeof s.index === 'number') setIndex(s.index);
      if (s.running === false) setPresent(false);
      else if (s.running === true) setPresent(true);
    });
    return () => off?.();
  }, [shell]);

  const save = useCallback(
    async (as = false) => {
      if (!doc) return false;
      let target = doc.path;
      if (as || !target) {
        target = await pickSave(shell, 'slides', doc.path || doc.name);
        // A cancelled Save As is not a save; the caller must know.
        if (!target) return false;
      }
      try {
        const saved = await shell.doc.save({ id: doc.id, path: target });
        setDoc((d) => ({ ...d, ...saved, dirty: false }));
        shell.app.addRecent({ path: saved.path, app: 'slides' }).catch(() => {});
        toast(`Saved ${saved.path.split(/[\\/]/).pop()}`, { tone: 'good' });
        return true;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return false;
      }
    },
    [doc, shell, toast]
  );

  // Closing a window with unsaved work must ask, not discard.
  useDirtyGuard({ shell, dirty: doc?.dirty, name: doc?.name, onSave: () => save(false) });
  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'slides');
    if (file) openInApp(shell, file, 'slides');
  }, [shell]);
  openFileRef.current = openFile;

  useFileDrop(useCallback((files) => files.forEach((f) => openInApp(shell, f, 'slides')), [shell]));

  /**
   * The Format Painter's second click: the armed look onto this shape — the
   * fill and outline into its own properties, the font onto every run of
   * its words — and the painter is put down, as PowerPoint's is after one.
   */
  const paintShape = useCallback(
    async (target) => {
      if (!painter) return;
      setPainter(null);
      if (!target || target.id === painter.from) return;
      const fill = painter.fill?.type === 'none' ? 'none' : painter.fill?.color ? painter.fill.color : null;
      const line = painter.line?.type === 'none' ? 'none' : painter.line?.color ? { color: painter.line.color, width: painter.line.width ?? 1, dash: painter.line.dash || null } : null;
      if (fill || line) await apply({ op: 'setShapeStyle', slide: index, shape: target.id, fill, line });
      if (painter.run && target.text) {
        const paragraphs = (target.text.paragraphs || []).map((p) => {
          const { runs, plain, ...props } = p;
          return { ...props, runs: (runs || []).map((r) => ({ ...r, ...painter.run, text: r.text })) };
        });
        await apply({ op: 'setText', slide: index, shape: target.id, paragraphs });
      }
      toast('Painted.', { ms: 1500 });
    },
    [apply, index, painter],
  );

  const commitText = useCallback(
    async (shapeId, text) => {
      // Each line keeps the paragraph it replaces — its level, bullet,
      // alignment and spacing — and the look of that paragraph's first run;
      // a line past the end takes the last paragraph's. Until 2026-09-21 the
      // words came back plain: a bulleted list edited once lost its bullets
      // and every bold word.
      const before = model?.slide?.shapes?.find((s) => s.id === shapeId)?.text?.paragraphs || [];
      const paragraphs = String(text).split('\n').map((line, i) => {
        const src = before[Math.min(i, before.length - 1)] || {};
        const { runs, plain, ...props } = src;
        const look = (runs || []).find((r) => r.text && r.text !== '\n') || {};
        const { text: _text, field: _field, break: _break, link: _link, ...runProps } = look;
        return { ...props, runs: [{ ...runProps, text: line }] };
      });
      await apply({ op: 'setText', slide: index, shape: shapeId, paragraphs });
      setEditing(null);
    },
    [apply, index, model]
  );

  /**
   * A slide that did not exist before, after the one on screen, and selected.
   *
   * Selecting it matters: a new slide you then have to go and find is a new
   * slide you did not want yet.
   */
  const addSlide = useCallback(
    async (layout = 'obj') => {
      const next = await apply({
        op: 'insertSlide',
        after: index,
        layout,
        ...(layout === 'blank' ? {} : { title: 'New slide', body: layout === 'title' ? '' : ['Point one'] }),
      });
      if (next) setIndex(Math.min(index + 1, (next.model?.count || index + 2) - 1));
      return next;
    },
    [apply, index]
  );
  const addSlideRef = useRef(null);
  addSlideRef.current = addSlide;

  /**
   * Start the show with a presenter window beside it.
   *
   * The second window opens on the same document — the session lives in the
   * main process, so both windows are looking at one deck rather than at two
   * copies of it — and this window goes full screen showing only the slide.
   */
  const presentWithNotes = useCallback(async () => {
    if (!doc) return;
    await shell.present.set({ id: doc.id, index, running: true, blank: false });
    await shell.win.create({ app: 'slides', query: { presenter: doc.id } });
    setPresent(true);
    shell.win.fullscreen({ on: true });
  }, [shell, doc, index]);

  const exportAs = useCallback(
    async (format, options = null) => {
      if (!doc) return;
      const target = await shell.dialog.save({
        title: `Export as ${format.toUpperCase()}`,
        defaultPath: (doc.path || doc.name).replace(/\.[^.]+$/, `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format] }],
      });
      if (!target) return;
      try {
        // A PDF is a print job: the page setup decides where the pages fall,
        // and a workbook or a deck must be laid out before it is one.
        if (format === 'pdf') await shell.print.pdf({ id: doc.id, path: target, options: options || defaultPrintOptions('deck') });
        else await shell.doc.export({ id: doc.id, format, path: target });
        toast(`Exported ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [doc, shell, toast]
  );

  const commands = useMemo(
    () => ({
      'file.new': { label: 'New', icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'slides' }) },
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: 'Save', icon: 'save', key: 'Mod+S', run: () => save(false) },
      'file.print': { label: 'Print…', icon: 'print', key: 'Mod+P', global: true, run: () => setPrinting(true) },
      'edit.find': { label: 'Find and replace…', icon: 'find', key: 'Mod+F', global: true, run: () => setFindOpen(true) },
      'insert.link': { label: 'Link…', icon: 'link', key: 'Mod+K', run: () => act('link') },
      'slide.next': { label: 'Next slide', icon: 'chevronRight', key: 'arrowdown', run: () => setIndex((i) => Math.min(i + 1, (model?.count || 1) - 1)) },
      'slide.prev': { label: 'Previous slide', icon: 'chevronLeft', key: 'arrowup', run: () => setIndex((i) => Math.max(0, i - 1)) },
      'slide.new': { label: 'Duplicate slide', icon: 'plus', run: () => apply({ op: 'duplicateSlide', slide: index }) },
      'slide.delete': { label: 'Delete slide', icon: 'trash', run: () => apply({ op: 'removeSlide', slide: index }) },
      'slide.textbox': { label: 'Text box', icon: 'textbox', run: () => apply({ op: 'addTextBox', slide: index, x: 120, y: 120, w: 420, h: 90, paragraphs: [{ runs: [{ text: 'New text' }] }] }) },
      'slide.picture': { label: 'Picture…', icon: 'picture', run: insertPicture },

      'view.present': { label: 'Present', icon: 'play', key: 'F5', run: () => setPresent(true) },
      'slide.add': { label: 'New slide', icon: 'plus', key: 'Mod+M', run: () => addSlideRef.current?.('obj') },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, model, index, apply, save, openFile, shell, insertPicture]

  );

  useCommands(commands, [doc, model, index]);

  useEffect(() => {
    if (!present) return undefined;
    const onKey = (e) => {
      if (e.key === 'Escape') setPresent(false);
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') setIndex((i) => Math.min(i + 1, (model?.count || 1) - 1));
      if (e.key === 'ArrowLeft' || e.key === 'PageUp') setIndex((i) => Math.max(0, i - 1));
    };
    window.addEventListener('keydown', onKey);
    if (!reading) shell.win.fullscreen({ on: true });
    return () => {
      window.removeEventListener('keydown', onKey);
      if (!reading) shell.win.fullscreen({ on: false });
      setReading(false);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [present, model, shell]);

  // A new slide means a new selection: the shape ids belong to the slide.
  useEffect(() => { setSelected(pendingSelect.current ?? null); pendingSelect.current = null; setPainter(null); }, [index]);


  const slide = model?.slide;
  const selectedShape = selected ? slide?.shapes?.find((s) => s.id === selected) || null : null;
  const scale = view.zoom ?? fit;
  dragRef.current = { scale };
  // What the ribbon shows for the selected shape: its first run's look and
  // its first paragraph's alignment — the granularity the writer edits at.
  const format = useMemo(() => {
    const p = selectedShape?.text?.paragraphs?.[0];
    const r = p?.runs?.[0] || {};
    return {
      bold: Boolean(r.bold), italic: Boolean(r.italic), underline: Boolean(r.underline), size: r.size || 18, color: r.color || null, font: r.font || '', align: p?.align || 'left',
      strike: Boolean(r.strike), spacing: r.spacing || 0, highlight: r.highlight || null,
      // The box's own: where the words sit, which way they run, how many columns.
      anchor: selectedShape?.text?.anchor || 'top', vert: selectedShape?.text?.vert || 'horz', columns: selectedShape?.text?.columns || 1,
      // The first paragraph's own list look: the bullet kind, its level and its line spacing.
      bullet: p?.bullet?.type || null, level: p?.level || 0, lineHeight: p?.lineHeight || null,
    };
  }, [selectedShape]);

  /**
   * The ribbon's verbs beyond one engine operation: the show, the view
   * modes and overlays, zoom and tone, windows, help — and the formatting
   * of the selected shape, applied to every run in it, which is the
   * granularity the deck writer edits at.
   */
  const act = async (name, arg) => {
    switch (name) {
      case 'present':
        if (arg === 'start') setIndex(0);
        if (arg === 'reading') setReading(true);
        setPresent(true);
        return;
      case 'mode': patchView({ mode: arg }); return;
      case 'toggle': patchView((v) => ({ [arg]: arg === 'notes' ? v.notes === false : !v[arg] })); return;
      case 'zoom': patchView({ zoom: arg }); return;
      case 'tone': patchView({ tone: arg }); return;
      // The right-hand pane: Layers (the slide's shapes, in drawing order)
      // or Designs (the deck's layouts). Asking for the one that is open closes it.
      case 'pane': patchView((v) => ({ pane: v.pane === arg ? null : arg })); return;
      // The Format pane opens (and stays open) from Shape Fill and Shape Outline.
      case 'formatPane': patchView({ pane: 'format' }); return;
      case 'dragEnd':
        // The box the pointer left the shape at.
        await apply({ op: 'setGeometry', slide: index, shape: arg.id, x: Math.round(arg.g.x), y: Math.round(arg.g.y), w: Math.round(arg.g.w), h: Math.round(arg.g.h) });
        return;
      case 'shapeFill':
        if (!selectedShape) return;
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, fill: arg });
        return;
      case 'shapeLine': {
        // A change to one of colour, weight or dashes keeps the other two.
        if (!selectedShape) return;
        const cur = selectedShape.line && selectedShape.line.type !== 'none' ? selectedShape.line : null;
        const line = arg === 'none'
          ? 'none'
          : { color: arg.color ?? cur?.color ?? (slide?.theme?.colors?.accent1 || '#4472C4'), width: arg.width ?? cur?.width ?? 1, dash: arg.dash ?? cur?.dash ?? null };
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, line });
        return;
      }
      case 'shapeShadow': {
        if (!selectedShape) return;
        const preset = SHADOW_PRESETS[arg];
        if (preset === undefined) return;
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, effects: preset });
        return;
      }
      case 'quickStyle':
        // Filled in an accent, outlined in the same accent darkened — the theme's own look.
        if (!selectedShape) return;
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, fill: { scheme: `accent${arg}` }, line: { color: { scheme: `accent${arg}`, lumMod: 50 }, width: 1 } });
        return;
      case 'order':
        if (!selectedShape) return;
        await apply({ op: 'reorderShape', slide: index, shape: selectedShape.id, to: arg });
        return;
      case 'hideShape':
        await apply({ op: 'setShapeHidden', slide: index, shape: arg.id, hidden: arg.hidden });
        return;
      case 'renameShape':
        await apply({ op: 'renameShape', slide: index, shape: arg.id, name: arg.name });
        return;
      case 'resetSlide':
        await apply({ op: 'resetSlide', slide: index });
        return;
      case 'body': {
        if (!selectedShape?.text) return toast('Click a text box first.', { ms: 3500 });
        await apply({ op: 'setBodyProps', slide: index, shape: selectedShape.id, ...arg });
        return;
      }
      case 'applyLayout':
        await apply({ op: 'applyLayout', slide: index, layout: arg });
        return;
      case 'newSlideFrom': {
        // A slide after this one on the chosen layout, with the words a new
        // slide of that kind starts with.
        const kind = arg.type === 'title' ? 'title' : arg.type === 'blank' || !arg.placeholders?.length ? 'blank' : 'obj';
        const next = await apply({ op: 'insertSlide', after: index, layout: kind, layoutPart: arg.part, ...(kind === 'blank' ? {} : { title: 'New slide', body: kind === 'title' ? '' : ['Point one'] }) });
        if (next) setIndex(Math.min(index + 1, (next.model?.count || index + 2) - 1));
        return;
      }
      case 'newWindow':
        if (!doc?.path) return toast('Save the presentation first, so a second window can open the same file.', { ms: 5000 });
        shell.win.create({ app: 'slides', file: doc.path });
        return;
      case 'help': shell.shell.openExternal({ url: SITE.help }); return;
      case 'feedback': shell.shell.openExternal({ url: SITE.contact }); return;
      case 'releases': shell.shell.openExternal({ url: SITE.releases }); return;
      case 'shortcuts': setShortcutsOpen(true); return;
      case 'addShape': {
        // Placed in the middle of the slide at a hand's size, then selected so
        // Arrange and the font controls act on it at once.
        const W = model?.size?.width || 1280;
        const H = model?.size?.height || 720;
        const w = arg === 'line' ? 320 : 240;
        const h = arg === 'line' ? 0 : 160;
        const next = await apply({ op: 'addShape', slide: index, preset: arg, x: Math.round((W - w) / 2), y: Math.round((H - h) / 2), w, h });
        const added = next?.model?.slide?.shapes?.slice(-1)[0];
        if (added) setSelected(added.id);
        return;
      }
      case 'footer': {
        setFooterOpen(arg || 'open');
        return;
      }
      case 'find': {
        setFindOpen(true);
        return;
      }
      case 'link': {
        if (!selectedShape?.text) return toast('Click a text box first.', { ms: 3000 });
        setLinkOpen(true);
        return;
      }
      // Sections: PowerPoint's Home → Section menu. A new one starts at this
      // slide and is named at once; Rename and Remove act on this slide's.
      case 'addSection': {
        const next = await apply({ op: 'addSection', slide: index, name: 'Untitled Section' });
        const at = (next?.model?.sections || []).findIndex((s) => s.slides.includes(index));
        if (at >= 0) setSectionRename({ section: at, name: next.model.sections[at].name });
        return;
      }
      case 'renameSection': {
        const own = (model?.sections || []).find((s) => s.slides.includes(index));
        if (!own) return toast('This slide is in no section. Add Section starts one at it.', { ms: 3500 });
        setSectionRename({ section: own.index, name: own.name });
        return;
      }
      case 'removeSection': {
        const own = (model?.sections || []).find((s) => s.slides.includes(index));
        if (!own) return toast('This slide is in no section.', { ms: 3000 });
        await apply({ op: 'removeSection', section: own.index });
        return;
      }
      case 'removeAllSections':
        await apply({ op: 'removeAllSections' });
        return;
      case 'clearFormat': {
        if (!selectedShape?.text) return toast('Click a text box first.', { ms: 3000 });
        await apply({ op: 'clearTextFormat', slide: index, shape: selectedShape.id });
        return;
      }
      case 'painter': {
        if (painter) { setPainter(null); toast('Format Painter put down.', { ms: 2000 }); return; }
        if (!selectedShape) return toast('Click a shape first.', { ms: 3000 });
        const look = (selectedShape.text?.paragraphs || []).flatMap((p) => p.runs || []).find((r) => r.text && r.text !== '\n') || null;
        const { text: _text, field: _field, break: _break, link: _link, ...runProps } = look || {};
        setPainter({ from: selectedShape.id, fill: selectedShape.fill ?? null, line: selectedShape.line ?? null, run: look ? runProps : null });
        toast('Format Painter: click a shape to give it this look. Esc puts it down.', { ms: 4000 });
        return;
      }
      case 'copyShape': {
        if (!selectedShape) return toast('Click a shape first.', { ms: 3000 });
        const got = await shell.doc.shapeClip({ id: doc.id, slide: index, shape: selectedShape.id });
        if (!got) return;
        setClip({ ...got, geometry: selectedShape.geometry });
        toast('Copied. Paste puts it on the slide on screen, a little right and down.', { ms: 3000 });
        return;
      }
      case 'cutShape': {
        if (!selectedShape) return;
        const got = await shell.doc.shapeClip({ id: doc.id, slide: index, shape: selectedShape.id });
        if (!got) return;
        setClip({ ...got, geometry: selectedShape.geometry });
        await apply({ op: 'removeShape', slide: index, shape: selectedShape.id });
        setSelected(null);
        return;
      }
      case 'pasteShape': {
        if (!clip) return toast('Nothing copied yet: click a shape and press Copy or Ctrl+C.', { ms: 3500 });
        const g = clip.geometry || { x: 100, y: 100, w: 300, h: 80 };
        await apply({ op: 'pasteShape', slide: index, clip: { xml: clip.xml, tag: clip.tag, rels: clip.rels }, geometry: { x: g.x + 20, y: g.y + 20, w: g.w, h: g.h, rot: g.rot } });
        // The pasted shape is the last in the drawing order: select it, so a
        // second paste or a nudge acts on it.
        const fresh = await shell.doc.model({ id: doc.id, slide: index }).catch(() => null);
        const shapes = fresh?.slide?.shapes || [];
        if (shapes.length) setSelected(shapes[shapes.length - 1].id);
        return;
      }
      case 'deleteShape':
        if (!selectedShape) return;

        await apply({ op: 'removeShape', slide: index, shape: selectedShape.id });
        setSelected(null);
        return;
      case 'nudge': {
        if (!selectedShape?.geometry) return;
        const g = selectedShape.geometry;
        await apply({ op: 'setGeometry', slide: index, shape: selectedShape.id, x: g.x + (arg.dx || 0), y: g.y + (arg.dy || 0), w: g.w, h: g.h });
        return;
      }
      case 'format': {
        if (!selectedShape?.text) return toast('Click a text box first.', { ms: 3500 });
        // Change case rewrites the letters, as PowerPoint's does: sentence
        // case from the paragraph's first run, the rest per run.
        const recase = (s, mode, first) => mode === 'upper' ? s.toUpperCase()
          : mode === 'lower' ? s.toLowerCase()
          : mode === 'title' ? s.replace(/(^|\s)(\S)/g, (m, sp, ch) => sp + ch.toUpperCase())
          : mode === 'sentence' ? (first ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s.toLowerCase())
          : s;
        const paragraphs = selectedShape.text.paragraphs.map((p) => {
          const { plain, ...rest } = p;
          // Bullets, numbering, the list level (a step from the paragraph's
          // own) and line spacing ride on the paragraph; the rest is as read.
          const level = 'level' in arg ? Math.max(0, Math.min(8, (p.level || 0) + Number(arg.level || 0))) : p.level;
          const bullet = arg.bullet === 'char' ? { type: 'char', char: '•' }
            : arg.bullet === 'number' ? { type: 'number', scheme: 'arabicPeriod', start: 1 }
            : arg.bullet === 'none' ? { type: 'none' }
            : p.bullet;
          return {
            ...rest,
            level: level || undefined,
            bullet,
            lineHeight: 'lineHeight' in arg ? arg.lineHeight : p.lineHeight,
            align: arg.align ?? p.align,
            runs: (p.runs || []).map((r, ri) => ({
              ...r,
              text: arg.case && r.text && r.text !== '\n' ? recase(r.text, arg.case, ri === 0) : r.text,
              bold: arg.bold === 'toggle' ? !r.bold : arg.bold ?? r.bold,
              italic: arg.italic === 'toggle' ? !r.italic : arg.italic ?? r.italic,
              underline: arg.underline === 'toggle' ? !r.underline : arg.underline ?? r.underline,
              strike: arg.strike === 'toggle' ? !r.strike : arg.strike ?? r.strike,
              spacing: 'spacing' in arg ? (arg.spacing || undefined) : r.spacing,
              highlight: 'highlight' in arg ? (arg.highlight || undefined) : r.highlight,
              size: arg.size ?? r.size,
              color: arg.color ?? r.color,
              font: 'font' in arg ? (arg.font || undefined) : r.font,
            })),
          };
        });
        await apply({ op: 'setText', slide: index, shape: selectedShape.id, paragraphs });
        return;
      }
      default:
        toast(`${name} is not wired yet.`, { ms: 3000 });
    }
  };

  actRef.current = act;

  // Every hook above, every early return below. This return sat above the
  // formatting memo, so a file the engine refused made React throw
  // "rendered fewer hooks than expected" and the person got a blank window
  // instead of the reason — for a truncated deck, twenty seconds of nothing.
  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Presentation" menu={appMenu}>
        <Empty icon="slides" title="This file could not be opened">{error}</Empty>
      </AppFrame>
    );
  }

  // A presenter window draws only the speaker's side. It shares the document
  // session, so nothing is opened twice and nothing can drift.
  if (presenterFor) {

    return (
      <AppFrame app={app} shell={shell} title="Presenter view" menu={appMenu}>
        <Presenter shell={shell} docId={presenterFor} />
      </AppFrame>
    );
  }

  if (present && slide) {
    return (
      <div
        className="sl-present"
        onClick={() => (led ? shell.present.set({ index: Math.min(index + 1, (model.count || 1) - 1) }) : setIndex((i) => Math.min(i + 1, (model.count || 1) - 1)))}
      >
        <style>{CSS}</style>
        {/* A black screen is a thing speakers ask for by name: attention back on them. */}
        {blank ? null : <div className="sl-present-stage" dangerouslySetInnerHTML={{ __html: slide.svg }} />}
        <div className="sl-present-bar">
          {index + 1} / {model.count}{led ? ' · driven from the presenter window' : ' · press Esc to leave'}
        </div>
      </div>
    );
  }

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={doc?.name || 'Presentation'}
      subtitle={doc?.converted ? `from ${doc.converted.from.toUpperCase()}` : null}
      dirty={doc?.dirty}
      menu={appMenu}
      ribbon={
        <SlidesRibbon
          tab={tab}
          setTab={setTab}
          model={model}
          doc={doc}
          commands={commands}
          shell={shell}
          menu={menu}
          save={save}
          openFile={openFile}
          exportAs={exportAs}
          act={act}
          view={view}
          index={index}
          selected={selected}
          format={format}
          canPaste={Boolean(clip)}
          painter={Boolean(painter)}
          addSlide={addSlide}
          insertPicture={insertPicture}
          presentWithNotes={presentWithNotes}

          setPresent={setPresent}
          setNotesOpen={setNotesOpen}
        />
      }
      status={
        <>
          <span>{doc?.path || 'Not saved yet'}</span>
          <Spacer />
          <Chip>Slide {index + 1} of {model?.count ?? 0}</Chip>
          {slide?.shapes ? <Chip>{slide.shapes.length} shapes</Chip> : null}
          <ZoomSlider value={view.zoom ?? fit} min={0.25} max={3} onChange={(v) => act('zoom', v)} onReset={() => act('zoom', null)} resetLabel="Fit to window" />
        </>
      }
    >
      {busy || !model ? (
        <div style={{ flex: 1, display: 'grid', placeItems: 'center' }}>
          <Spinner style={{ width: 22, height: 22 }} />
        </div>
      ) : (
        <>
          <style>{CSS}</style>
          <Panel width={196} resizable title="Slides">
            <div className="sl-sorter">
              {(model.outline || []).map((o, i) => (
                <React.Fragment key={o.part || i}>
                  {sectionHeading(model, o, i, (s) => setSectionRename({ section: s.index, name: s.name }), () => setIndex(i))}
                  <button
                    type="button"
                    className={`sl-thumb${i === index ? ' active' : ''}`}
                    onClick={() => setIndex(i)}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['slide.new', 'slide.delete']))}
                  >
                    <span className="sl-thumb-n">{i + 1}</span>
                    <span className="sl-thumb-card" title={o.title || `Slide ${i + 1}`}>
                      {o.thumbnail
                        ? <span className="sl-thumb-pic" dangerouslySetInnerHTML={{ __html: o.thumbnail }} />
                        : <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>}
                    </span>
                  </button>
                </React.Fragment>
              ))}
              {emptySectionHeadings(model, (s) => setSectionRename({ section: s.index, name: s.name }))}
            </div>
          </Panel>

          <Content>
            <div
              className={`sl-stage tone-${view.tone || 'colour'}`}
              ref={stageRef}
              tabIndex={0}
              onMouseDown={(e) => { if (e.target === e.currentTarget) setSelected(null); }}
              onKeyDown={(e) => {
                // The keyboard on the stage: arrows nudge the selected shape
                // (a pixel, ten with Shift), Delete removes it, Escape lets go.
                if (!editing && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') { e.preventDefault(); act('pasteShape'); return; }
                if (editing || !selectedShape) return;
                if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') { e.preventDefault(); act('copyShape'); return; }
                if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') { e.preventDefault(); act('cutShape'); return; }
                const step = e.shiftKey ? 10 : 1;
                const nudge = { ArrowLeft: { dx: -step }, ArrowRight: { dx: step }, ArrowUp: { dy: -step }, ArrowDown: { dy: step } }[e.key];
                if (nudge) { e.preventDefault(); act('nudge', nudge); }
                else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); act('deleteShape'); }
                else if (e.key === 'Escape') { if (painter) setPainter(null); else setSelected(null); }
              }}
            >
              {slide && view.mode === 'sorter' ? (
                <div className="sl-sortergrid">
                  {(model.outline || []).map((o, i) => (
                    <React.Fragment key={o.part || i}>
                      {sectionHeading(model, o, i, (s) => setSectionRename({ section: s.index, name: s.name }), () => setIndex(i))}
                      <button type="button" className={`sl-sortercard${i === index ? ' active' : ''}`} onClick={() => { setIndex(i); patchView({ mode: 'normal' }); }} title={o.title || `Slide ${i + 1}`}>
                        {o.thumbnail ? <span className="sl-thumb-pic" dangerouslySetInnerHTML={{ __html: o.thumbnail }} /> : <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>}
                        <span className="sl-sortern">{i + 1}</span>
                      </button>
                    </React.Fragment>
                  ))}
                  {emptySectionHeadings(model, (s) => setSectionRename({ section: s.index, name: s.name }))}
                </div>
              ) : slide && view.mode === 'outline' ? (
                <div className="sl-outline">
                  {(model.outline || []).map((o, i) => (
                    <button key={o.part || i} type="button" className={`sl-outlineitem${i === index ? ' active' : ''}`} onClick={() => setIndex(i)}>
                      <span className="sl-sortern">{i + 1}</span>
                      <span className="grow">
                        <div className="sl-outlinetitle">{o.title || 'Untitled slide'}</div>
                        {i === index ? slide.shapes.filter((s) => s.text).map((s) => s.text.paragraphs.map((p) => p.plain).join(' ')).filter((t) => t && t !== o.title).map((t, k) => <div key={k} className="sl-outlinetext">{t}</div>) : null}
                      </span>
                    </button>
                  ))}
                </div>
              ) : slide ? (
                <div className="sl-fit" style={{ width: Math.round(model.size.width * (view.zoom ?? fit)), height: Math.round(model.size.height * (view.zoom ?? fit)) }}>
                {view.ruler ? <><div className="sl-ruler-h" /><div className="sl-ruler-v" /></> : null}
                <div className="sl-slide" style={{ width: model.size.width, height: model.size.height, transform: `scale(${view.zoom ?? fit})`, transformOrigin: 'top left' }}>
                  <div className="sl-svg" dangerouslySetInnerHTML={{ __html: slide.svg }} />
                  {view.gridlines ? <div className="sl-gridlines" /> : null}
                  {view.guides ? <div className="sl-guides" /> : null}
                  {/*
                    Every shape gets a hit area over the drawing, so a click lands
                    on the shape: click selects, drag moves, a double-click on
                    words edits them. The selected one wears eight handles.
                  */}
                  {slide.shapes.filter((s) => s.geometry && !s.hidden).map((s) => {
                    const g = drag?.id === s.id ? drag.g : s.geometry;
                    return (
                      <button
                        key={s.id}
                        type="button"
                        className={`sl-hit${selected === s.id ? ' selected' : ''}${drag?.id === s.id ? ' dragging' : ''}`}
                        data-shape={s.id}
                        style={{ left: g.x, top: g.y, width: g.w, height: g.h }}
                        onMouseDown={(e) => startDrag(e, s, 'move')}
                        onClick={(e) => {
                          if ((e.ctrlKey || e.metaKey) && linkOf(s)) { shell.shell.openExternal({ url: linkOf(s) }); return; }
                          if (painter) paintShape(s);
                          setSelected(s.id);
                        }}
                        onDoubleClick={() => (s.text ? setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') }) : null)}
                        onContextMenu={(e) => menu.open(e, [
                          ...(s.text ? [{ label: 'Edit text', icon: 'textbox', run: () => setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') }) }] : []),
                          { label: 'Format shape…', icon: 'wand', run: () => { setSelected(s.id); act('formatPane'); } },
                          { label: 'Bring to front', icon: 'chevronUp', run: () => { setSelected(s.id); apply({ op: 'reorderShape', slide: index, shape: s.id, to: 'front' }); } },
                          { label: 'Send to back', icon: 'chevronDown', run: () => { setSelected(s.id); apply({ op: 'reorderShape', slide: index, shape: s.id, to: 'back' }); } },
                          '-',
                          { label: 'Delete shape', icon: 'trash', run: () => apply({ op: 'removeShape', slide: index, shape: s.id }) },
                        ])}
                        title={(s.text ? `${s.name || 'Shape'} — drag to move, double-click to edit` : `${s.name || s.kind} — drag to move`) + (linkOf(s) ? ` — Ctrl+click to follow ${linkOf(s)}` : '')}
                      />
                    );
                  })}
                  {selectedShape?.geometry && !selectedShape.hidden && !editing
                    ? HANDLES.map(([name, fx, fy, cursor]) => {
                        const g = drag?.id === selectedShape.id ? drag.g : selectedShape.geometry;
                        const size = 9 / scale;
                        return (
                          <div
                            key={name}
                            className="sl-handle"
                            data-handle={name}
                            style={{ left: g.x + g.w * fx - size / 2, top: g.y + g.h * fy - size / 2, width: size, height: size, cursor, borderWidth: 1.5 / scale }}
                            onMouseDown={(e) => startDrag(e, selectedShape, 'resize', name)}
                          />
                        );
                      })
                    : null}
                  {editing ? (
                    <textarea
                      className="sl-editor"
                      autoFocus
                      defaultValue={editing.text}
                      style={(() => {
                        const s = slide.shapes.find((x) => x.id === editing.id);
                        return s?.geometry ? { left: s.geometry.x, top: s.geometry.y, width: s.geometry.w, height: s.geometry.h } : {};
                      })()}
                      onBlur={(e) => commitText(editing.id, e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Escape') setEditing(null);
                        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commitText(editing.id, e.target.value);
                      }}
                    />
                  ) : null}
                </div>
                </div>
              ) : (
                <Empty icon="slides" title="This presentation has no slides" />
              )}
            </div>
            {slide && view.mode === 'notes' ? (
              <textarea
                key={`notes-${index}`}
                className="sl-notespage"
                defaultValue={slide.notes || ''}
                placeholder="Click to add notes"
                onBlur={async (e) => { if (e.target.value !== (slide.notes || '')) await apply({ op: 'setNotes', slide: index, text: e.target.value }); }}
              />
            ) : slide?.notes && view.notes !== false ? <div className="sl-notes">{slide.notes}</div> : null}
          </Content>

          {view.pane && model ? (
            <Panel
              right
              width={252}
              resizable
              title={view.pane === 'layers' ? 'Layers' : view.pane === 'designs' ? 'Designs' : 'Format'}
              actions={<Button icon="close" title="Close the pane" onClick={() => act('pane', view.pane)} />}
            >
              {view.pane === 'layers' ? (
                <LayersPane slide={slide} selected={selected} onSelect={setSelected} act={act} />
              ) : view.pane === 'designs' ? (
                <DesignsPane layouts={model.layouts} current={slide?.layout || null} size={model.size} act={act} />
              ) : (
                <FormatPane shape={selectedShape} theme={slide?.theme} act={act} />
              )}
            </Panel>
          ) : null}
          {menu.node}
        </>
      )}

      {shortcutsOpen ? <SlidesShortcutsDialog onClose={() => setShortcutsOpen(false)} /> : null}

      {printing && doc ? (
        <PrintDialog
          shell={shell}
          doc={doc}
          kind="deck"
          onClose={() => setPrinting(false)}
          onSaveAs={(options) => exportAs('pdf', options)}
        />
      ) : null}

      {linkOpen && selectedShape ? (
        <ShapeLinkDialog
          current={linkOf(selectedShape)}
          words={(selectedShape.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join(' ').trim()}
          onClose={() => setLinkOpen(false)}
          onApply={async (url) => {
            await apply({ op: 'setLink', slide: index, shape: selectedShape.id, url });
            setLinkOpen(false);
          }}
          onRemove={async () => {
            await apply({ op: 'setLink', slide: index, shape: selectedShape.id, url: null });
            setLinkOpen(false);
          }}
        />
      ) : null}

      {sectionRename ? (
        <SectionNameDialog
          name={sectionRename.name}
          onClose={() => setSectionRename(null)}
          onApply={async (name) => {
            await apply({ op: 'renameSection', section: sectionRename.section, name });
            setSectionRename(null);
          }}
        />
      ) : null}

      {findOpen ? (
        <FindDialog
          onClose={() => setFindOpen(false)}
          onFind={async (find, matchCase) => {
            const count = model?.count || 1;
            const target = matchCase ? find : find.toLowerCase();
            const hits = [];
            for (let i = 0; i < count; i++) {
              const m = await shell.doc.model({ id: doc.id, slide: i }).catch(() => null);
              for (const s of m?.slide?.shapes || []) {
                const words = (s.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join(' ');
                const hay = matchCase ? words : words.toLowerCase();
                let n = 0;
                for (let at = hay.indexOf(target); at !== -1; at = hay.indexOf(target, at + target.length)) n += 1;
                if (n) hits.push({ slide: i, shape: s.id, name: s.name || s.kind, count: n, text: words.replace(/\s+/g, ' ').trim().slice(0, 80) });
              }
            }
            return hits;
          }}
          onGoto={(hit) => {
            if (hit.slide === index) setSelected(hit.shape);
            else { pendingSelect.current = hit.shape; setIndex(hit.slide); }
          }}
          onReplaceAll={async (find, replace, matchCase, expected) => {
            const next = await apply({ op: 'replaceText', find, replace, matchCase });
            return next ? `Replaced ${expected} across the deck.` : 'Nothing was replaced.';
          }}
        />
      ) : null}

      {footerOpen ? (
        <FooterDialog
          key={index}
          slide={index}
          shapes={slide?.shapes || []}
          preset={footerOpen}
          onClose={() => setFooterOpen(null)}
          onApply={async (spec, all) => {
            await apply({ op: 'setFooter', slide: index, all, ...spec });
            setFooterOpen(null);
            toast(all ? 'Footer applied to every slide' : 'Footer applied to this slide', { tone: 'good' });
          }}
        />
      ) : null}

      {notesOpen ? (
        <NotesDialog
          key={index}
          slide={index}
          text={slide?.notes || ''}
          onClose={() => setNotesOpen(false)}
          onSave={async (text) => {
            await apply({ op: 'setNotes', slide: index, text });
            setNotesOpen(false);
            toast('Notes saved', { tone: 'good' });
          }}
        />
      ) : null}
    </AppFrame>
  );
}

/** The eight handles: where each sits on the box, and the cursor it shows. */
const HANDLES = [
  ['nw', 0, 0, 'nwse-resize'], ['n', 0.5, 0, 'ns-resize'], ['ne', 1, 0, 'nesw-resize'],
  ['e', 1, 0.5, 'ew-resize'], ['se', 1, 1, 'nwse-resize'], ['s', 0.5, 1, 'ns-resize'],
  ['sw', 0, 1, 'nesw-resize'], ['w', 0, 0.5, 'ew-resize'],
];

/** The box a handle drags to, never smaller than a few pixels; `keep` holds the proportions on a corner. */
function resized(g0, handle, dx, dy, keep) {
  let { x, y, w, h } = g0;
  if (handle.includes('e')) w = g0.w + dx;
  if (handle.includes('s')) h = g0.h + dy;
  if (handle.includes('w')) { x = g0.x + dx; w = g0.w - dx; }
  if (handle.includes('n')) { y = g0.y + dy; h = g0.h - dy; }
  if (keep && handle.length === 2 && g0.w > 0 && g0.h > 0) {
    const ratio = g0.w / g0.h;
    if (Math.abs(w / g0.w) > Math.abs(h / g0.h)) h = w / ratio; else w = h * ratio;
    if (handle.includes('w')) x = g0.x + g0.w - w;
    if (handle.includes('n')) y = g0.y + g0.h - h;
  }
  if (w < 8) { if (handle.includes('w')) x = g0.x + g0.w - 8; w = 8; }
  if (h < 8) { if (handle.includes('n')) y = g0.y + g0.h - 8; h = 8; }
  return { x, y, w, h };
}

/** The theme's colours, in the order PowerPoint's gallery shows them, and its standard row. */
const THEME_SWATCHES = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
const STANDARD_SWATCHES = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
const LINE_DASHES = [['solid', 'Solid'], ['sysDash', 'Round dot'], ['dash', 'Dash'], ['dashDot', 'Dash dot'], ['lgDash', 'Long dash'], ['lgDashDot', 'Long dash dot']];

/**
 * The selected shape's fill and outline — PowerPoint's Format Shape pane,
 * with the parts people use: the theme's colours and the standard ones as
 * swatches, any colour, no fill; the outline's colour, weight and dashes,
 * or no outline. Each press is one engine operation on the shape's own
 * properties, which the file keeps.
 */
/** Shape Effects: PowerPoint's offset shadows, in points and degrees, and none. */
const SHADOW_PRESETS = {
  none: 'none',
  br: { shadow: { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 } },
  b: { shadow: { dist: 3, dir: 90, blur: 4, color: '#000000', alpha: 0.4 } },
  r: { shadow: { dist: 3, dir: 0, blur: 4, color: '#000000', alpha: 0.4 } },
  tl: { shadow: { dist: 3, dir: 225, blur: 4, color: '#000000', alpha: 0.4 } },
  c: { shadow: { dist: 0, dir: 0, blur: 6, color: '#000000', alpha: 0.45 } },
};
const SHADOW_LABELS = [['none', 'None'], ['br', 'Bottom right'], ['b', 'Below'], ['r', 'Right'], ['tl', 'Top left'], ['c', 'All round']];
/** Which preset a shape's shadow is, for the pane to mark. */
function shadowKeyOf(effects) {
  const sh = effects?.shadow;
  if (!sh) return 'none';
  if (Math.round(sh.distPx) === 0) return 'c';
  const dir = Math.round(sh.dir);
  return dir === 45 ? 'br' : dir === 90 ? 'b' : dir === 0 ? 'r' : dir === 225 ? 'tl' : null;
}

function FormatPane({ shape, theme, act }) {
  if (!shape) return <div className="sl-pane-empty">Click a shape on the slide to format it.</div>;
  if (shape.kind === 'table' || shape.kind === 'chart' || shape.kind === 'unsupported') return <div className="sl-pane-empty">A table or chart frame has no fill or outline of its own.</div>;
  const colours = theme?.colors || {};
  const fill = shape.fill?.type === 'solid' ? shape.fill.color : shape.fill?.type === 'none' ? 'none' : null;
  const line = shape.line && shape.line.type !== 'none' ? shape.line : null;
  const same = (a, b) => Boolean(a && b && String(a).toLowerCase() === String(b).toLowerCase());
  const swatches = (onPick, current) => (
    <>
      <div className="sl-swatches">
        {THEME_SWATCHES.filter((n) => colours[n]).map((n) => (
          <button key={n} type="button" className={`sl-swatch${same(current, colours[n]) ? ' current' : ''}`} title={n} data-swatch={n} style={{ background: colours[n] }} onClick={() => onPick(colours[n])} />
        ))}
      </div>
      <div className="sl-swatches">
        {STANDARD_SWATCHES.map((hex) => (
          <button key={hex} type="button" className={`sl-swatch${same(current, hex) ? ' current' : ''}`} title={hex} data-swatch={hex} style={{ background: hex }} onClick={() => onPick(hex)} />
        ))}
      </div>
    </>
  );
  return (
    <div className="sl-format">
      <div className="sl-format-head">{shape.name || shape.kind}</div>
      <section className="sl-format-fill">
        <h4>Fill</h4>
        {swatches((hex) => act('shapeFill', hex), fill)}
        <div className="sl-format-row">
          <button type="button" className={`sl-chip${fill === 'none' ? ' current' : ''}`} onClick={() => act('shapeFill', 'none')}>No fill</button>
          <input type="color" className="sl-colour" title="Any colour" value={fill && fill !== 'none' ? fill : '#4472c4'} onChange={(e) => act('shapeFill', e.target.value)} />
        </div>
      </section>
      <section className="sl-format-line">
        <h4>Outline</h4>
        {swatches((hex) => act('shapeLine', { color: hex }), line?.color)}
        <div className="sl-format-row">
          <button type="button" className={`sl-chip${shape.line?.type === 'none' ? ' current' : ''}`} onClick={() => act('shapeLine', 'none')}>No outline</button>
          <input type="color" className="sl-colour" title="Any colour" value={line?.color || '#1f2937'} onChange={(e) => act('shapeLine', { color: e.target.value })} />
        </div>
        <div className="sl-format-row">
          <label>
            Weight
            <select className="rw-input" value={String(line?.width ?? 1)} onChange={(e) => act('shapeLine', { width: Number(e.target.value) })}>
              {LINE_WEIGHTS.map((w) => <option key={w} value={String(w)}>{w} pt</option>)}
            </select>
          </label>
          <label>
            Dashes
            <select className="rw-input" value={line?.dash || 'solid'} onChange={(e) => act('shapeLine', { dash: e.target.value })}>
              {LINE_DASHES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          </label>
        </div>
      </section>
      <section className="sl-format-shadow">
        <h4>Shadow</h4>
        <div className="sl-format-row" style={{ flexWrap: 'wrap' }}>
          {SHADOW_LABELS.map(([key, label]) => (
            <button key={key} type="button" className={`sl-chip sl-shadow-${key}${shadowKeyOf(shape.effects) === key ? ' current' : ''}`} onClick={() => act('shapeShadow', key)}>{label}</button>
          ))}
        </div>
      </section>
    </div>
  );
}

/** The icon a layer row shows for its shape. */
const LAYER_ICONS = { picture: 'picture', table: 'table', chart: 'chart', connector: 'minus', shape: 'shape', unsupported: 'shape' };

/**
 * The slide's shapes as layers, top-most first — PowerPoint's selection pane
 * with the parts people use: a click selects the shape on the stage, the eye
 * hides it (it stays in the file, undrawn), the arrows change the drawing
 * order, a double-click renames it.
 */
function LayersPane({ slide, selected, onSelect, act }) {
  const [renaming, setRenaming] = React.useState(null);
  const shapes = slide?.shapes || [];
  const rows = [...shapes].reverse();
  const pos = selected != null ? shapes.findIndex((s) => s.id === selected) : -1;
  const n = shapes.length;
  return (
    <div className="sl-layers">
      <div className="sl-layers-tools">
        <Button icon="chevronUp" title="Bring forward" disabled={pos < 0 || pos >= n - 1} onClick={() => act('order', 'forward')} />
        <Button icon="chevronDown" title="Send backward" disabled={pos <= 0} onClick={() => act('order', 'backward')} />
        <Button label="Front" title="Bring to front" disabled={pos < 0 || pos >= n - 1} onClick={() => act('order', 'front')} />
        <Button label="Back" title="Send to back" disabled={pos <= 0} onClick={() => act('order', 'back')} />
        <Spacer />
        <Button icon="trash" title="Delete the selected shape" disabled={pos < 0} onClick={() => act('deleteShape')} />
      </div>
      {rows.length ? (
        rows.map((s) => (
          <div
            key={s.id}
            className={`sl-layer${selected === s.id ? ' active' : ''}${s.hidden ? ' off' : ''}`}
            data-shape={s.id}
            onClick={() => onSelect(s.id)}
            onDoubleClick={() => setRenaming({ id: s.id, name: s.name || '' })}
            title="Click to select; double-click to rename"
          >
            <button
              type="button"
              className="sl-eye"
              title={s.hidden ? 'Show this shape' : 'Hide this shape'}
              onClick={(e) => {
                e.stopPropagation();
                act('hideShape', { id: s.id, hidden: !s.hidden });
              }}
            >
              <Icon name="eye" size={14} />
            </button>
            <Icon name={s.text ? 'textbox' : LAYER_ICONS[s.kind] || 'shape'} size={14} />
            {renaming?.id === s.id ? (
              <input
                autoFocus
                className="rw-input sl-layer-name"
                value={renaming.name}
                onChange={(e) => setRenaming({ ...renaming, name: e.target.value })}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    act('renameShape', { id: s.id, name: renaming.name });
                    setRenaming(null);
                  }
                  if (e.key === 'Escape') setRenaming(null);
                }}
                onBlur={() => setRenaming(null)}
                onClick={(e) => e.stopPropagation()}
              />
            ) : (
              <span className="sl-layer-text">
                <span className="sl-layer-title">{s.name || `${s.kind} ${s.id}`}</span>
                {s.text?.paragraphs?.length ? <span className="sl-layer-words">{s.text.paragraphs.map((p) => p.plain).join(' ').slice(0, 70)}</span> : null}
              </span>
            )}
          </div>
        ))
      ) : (
        <div className="sl-pane-empty">Nothing on this slide yet.</div>
      )}
    </div>
  );
}

/**
 * The deck's own layouts as a gallery — each drawn as the boxes its
 * placeholders make, the way PowerPoint's Layout gallery draws them. A
 * click puts the current slide on that layout; New starts a slide from it.
 */
function DesignsPane({ layouts, current, size, act }) {
  const W = size?.width || 960;
  const H = size?.height || 540;
  if (!layouts?.length) return <div className="sl-pane-empty">This deck has no layouts of its own.</div>;
  return (
    <div className="sl-designs">
      {layouts.map((l) => (
        <div
          key={l.part}
          className={`sl-design${l.part === current ? ' active' : ''}`}
          data-layout={l.part}
          title={l.part === current ? `${l.name} — this slide's layout` : `${l.name} — click to put this slide on it`}
          onClick={() => (l.part === current ? null : act('applyLayout', l.part))}
        >
          <svg className="sl-design-pic" viewBox={`0 0 ${W} ${H}`} aria-hidden="true">
            <rect x="0" y="0" width={W} height={H} className="sl-design-bg" />
            {l.placeholders.map((p, i) => (
              <rect key={i} x={p.geometry.x} y={p.geometry.y} width={Math.max(1, p.geometry.w)} height={Math.max(1, p.geometry.h)} rx={W * 0.006} className={`sl-design-ph${/title/i.test(p.type) ? ' title' : ''}`} />
            ))}
          </svg>
          <div className="sl-design-foot">
            <span className="sl-design-name">{l.name}</span>
            <button
              type="button"
              className="sl-design-new"
              title={`A new slide on the ${l.name} layout, after this one`}
              onClick={(e) => {
                e.stopPropagation();
                act('newSlideFrom', l);
              }}
            >
              + New
            </button>
          </div>
        </div>
      ))}
    </div>
  );
}

const CSS = `
/* the right-hand panes ------------------------------------------------------ */
.sl-format-head { padding: 8px 12px; font-weight: 600; font-size: 12.5px; border-bottom: 1px solid var(--line-soft); }
.sl-format section { padding: 10px 12px; border-bottom: 1px solid var(--line-soft); }
.sl-format h4 { margin: 0 0 8px; font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-3); font-weight: 600; }
.sl-swatches { display: grid; grid-template-columns: repeat(10, 1fr); gap: 4px; margin-bottom: 5px; }
.sl-swatch { height: 18px; border-radius: 4px; border: 1px solid rgba(0, 0, 0, .14); cursor: pointer; padding: 0; }
.sl-swatch:hover { transform: scale(1.12); }
.sl-swatch.current { outline: 2px solid var(--accent); outline-offset: 1px; }
.sl-format-row { display: flex; gap: 8px; align-items: center; margin-top: 6px; font-size: 12px; }
.sl-format-row label { display: flex; flex-direction: column; gap: 3px; flex: 1; font-size: 11.5px; color: var(--ink-2); }
.sl-format-row select.rw-input { height: 26px; font-size: 12px; }
.sl-chip { border: 1px solid var(--line); background: var(--surface); border-radius: 999px; padding: 3px 10px; font-size: 11.5px; color: var(--ink-2); }
.sl-chip:hover, .sl-chip.current { color: var(--accent); border-color: var(--accent); }
.sl-colour { width: 30px; height: 24px; padding: 0 2px; border: 1px solid var(--line); border-radius: 4px; background: var(--surface); }
.sl-pane-empty { padding: 14px; color: var(--ink-3); font-size: 12.5px; }
.sl-layers { display: flex; flex-direction: column; }
.sl-layers-tools { display: flex; align-items: center; gap: 2px; padding: 6px 8px; border-bottom: 1px solid var(--line-soft); }
.sl-layer { display: flex; align-items: center; gap: 8px; padding: 6px 10px; cursor: default; border-bottom: 1px solid var(--line-soft); font-size: 12.5px; color: var(--ink); }
.sl-layer:hover { background: var(--surface-2); }
.sl-layer.active { background: var(--selected); box-shadow: inset 3px 0 0 var(--accent); }
.sl-layer.off .sl-layer-text, .sl-layer.off > svg { opacity: .4; }
.sl-layer.off .sl-eye { opacity: .35; }
.sl-layer-text { display: flex; flex-direction: column; min-width: 0; flex: 1; }
.sl-layer-title { font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-layer-words { color: var(--ink-3); font-size: 11.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-eye { border: 0; background: transparent; color: var(--ink-3); padding: 2px; display: grid; place-items: center; border-radius: 4px; }
.sl-eye:hover { color: var(--ink); background: var(--surface-2); }
.sl-layer-name { flex: 1; min-width: 0; height: 24px; }
.sl-designs { display: grid; grid-template-columns: 1fr; gap: 10px; padding: 10px; }
.sl-design { border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; background: var(--surface); cursor: pointer; transition: box-shadow var(--fast), border-color var(--fast); }
.sl-design:hover { border-color: var(--accent); box-shadow: var(--shadow-1); }
.sl-design.active { border-color: var(--accent); box-shadow: 0 0 0 2px color-mix(in srgb, var(--accent) 30%, transparent); cursor: default; }
.sl-design-pic { display: block; width: 100%; height: auto; }
.sl-design-bg { fill: #fff; }
.sl-design-ph { fill: rgba(120, 130, 150, .18); stroke: rgba(120, 130, 150, .6); stroke-width: 6; stroke-dasharray: 18 12; }
.sl-design-ph.title { fill: color-mix(in srgb, var(--accent) 16%, transparent); stroke: color-mix(in srgb, var(--accent) 70%, transparent); }
.sl-design-foot { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 6px 9px; font-size: 12px; border-top: 1px solid var(--line-soft); }
.sl-design-name { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-design-new { flex: none; border: 1px solid var(--line); background: var(--surface); border-radius: 999px; padding: 2px 9px; font-size: 11.5px; color: var(--ink-2); }
.sl-design-new:hover { color: var(--accent); border-color: var(--accent); }

.sl-sorter { display: flex; flex-direction: column; gap: 6px; padding: 8px; }
/* A section's heading, above its first slide, as PowerPoint draws one: the name and how many slides. */
.sl-section { display: flex; align-items: baseline; gap: 6px; min-width: 0; padding: 6px 2px 0 20px; border: 0; background: transparent; text-align: left; cursor: pointer; }
.sl-sortergrid .sl-section { grid-column: 1 / -1; padding-left: 2px; }
.sl-section + .sl-section, .sl-sorter .sl-section:first-child { padding-top: 0; }
.sl-section-name { font-size: 11.5px; font-weight: 600; color: var(--ink-2); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-section:hover .sl-section-name { color: var(--accent); }
.sl-section-count { font-size: 10.5px; color: var(--ink-3); flex: none; }
.sl-thumb { display: flex; align-items: stretch; gap: 7px; border: 0; background: transparent; padding: 0; text-align: left; }
.sl-thumb-n { width: 16px; font-size: 11px; color: var(--ink-3); padding-top: 3px; flex: none; text-align: right; }
.sl-thumb-card {
  flex: 1; min-width: 0; aspect-ratio: 16 / 9; border: 2px solid var(--line); border-radius: var(--r-2);
  background: var(--surface); padding: 7px 8px; overflow: hidden; transition: border-color var(--fast);
}
.sl-thumb:hover .sl-thumb-card { border-color: var(--line-strong); }
.sl-thumb.active .sl-thumb-card { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.sl-thumb-title {
  font-size: 10.5px; line-height: 1.3; color: var(--ink-2); display: -webkit-box;
  -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden;
}
/* A real picture of the slide, the same drawing smaller. */
.sl-thumb-card:has(.sl-thumb-pic) { padding: 0; background: #fff; }
.sl-thumb-pic { display: block; line-height: 0; }
.sl-thumb-pic svg { display: block; width: 100%; height: auto; }
.sl-fit { position: relative; flex: none; }

.sl-stage { flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center; padding: 22px; background: var(--window); }
.sl-slide { position: relative; box-shadow: var(--shadow-2); background: #fff; }
.sl-svg svg { display: block; width: 100%; height: 100%; }
.sl-hit { position: absolute; border: 1px solid transparent; background: transparent; border-radius: 2px; min-height: 8px; min-width: 8px; cursor: move; }
.sl-hit.dragging { border: 1px dashed var(--accent); background: rgba(43, 95, 217, 0.08); z-index: 4; }
.sl-handle { position: absolute; z-index: 5; background: #fff; border: 1.5px solid var(--accent); border-radius: 2px; box-sizing: border-box; }
.sl-stage:focus { outline: none; }

.sl-hit:hover { border-color: var(--accent-line); background: rgba(43, 95, 217, 0.06); }
.sl-editor {
  position: absolute; border: 2px solid var(--accent); border-radius: 3px; padding: 4px 6px;
  font: inherit; font-size: 15px; background: #fff; color: #111; resize: none; outline: none; z-index: 5;
}
.sl-notes {
  border-top: 1px solid var(--line); background: var(--chrome); padding: 9px 18px;
  font-size: 12.5px; color: var(--ink-2); max-height: 110px; overflow: auto; white-space: pre-wrap;
}

.sl-hit.selected { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
/* View → Show: rulers beside the slide, gridlines and guides over it. */
.sl-ruler-h { position: absolute; left: 0; right: 0; top: -14px; height: 12px; background: repeating-linear-gradient(to right, var(--ink-3) 0 1px, transparent 1px 48px); opacity: .5; }
.sl-ruler-v { position: absolute; top: 0; bottom: 0; left: -14px; width: 12px; background: repeating-linear-gradient(to bottom, var(--ink-3) 0 1px, transparent 1px 48px); opacity: .5; }
.sl-gridlines { position: absolute; inset: 0; pointer-events: none; background-image: linear-gradient(to right, rgba(0,0,0,.12) 1px, transparent 1px), linear-gradient(to bottom, rgba(0,0,0,.12) 1px, transparent 1px); background-size: 48px 48px; }
.sl-guides { position: absolute; inset: 0; pointer-events: none; background-image: linear-gradient(to right, transparent calc(50% - 1px), rgba(200,0,0,.6) calc(50% - 1px), rgba(200,0,0,.6) 50%, transparent 50%), linear-gradient(to bottom, transparent calc(50% - 1px), rgba(200,0,0,.6) calc(50% - 1px), rgba(200,0,0,.6) 50%, transparent 50%); }
/* View → Colour/Greyscale. */
.sl-stage.tone-grey .sl-fit { filter: grayscale(1); }
.sl-stage.tone-mono .sl-fit { filter: grayscale(1) contrast(4); }
/* Slide Sorter and Outline View. */
.sl-sortergrid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 14px; width: 100%; align-self: start; }
.sl-sortercard { position: relative; border: 2px solid var(--line); border-radius: var(--r-2); background: #fff; padding: 0; overflow: hidden; aspect-ratio: 16 / 9; cursor: pointer; }
.sl-sortercard:hover { border-color: var(--line-strong); }
.sl-sortercard.active { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.sl-sortern { position: absolute; left: 6px; bottom: 4px; font-size: 11px; color: var(--ink-3); background: rgba(255,255,255,.85); padding: 0 5px; border-radius: 3px; }
.sl-outline { width: min(760px, 100%); align-self: start; display: flex; flex-direction: column; gap: 4px; }
.sl-outlineitem { display: flex; gap: 10px; align-items: flex-start; text-align: left; border: 1px solid transparent; border-radius: var(--r-2); background: transparent; padding: 6px 10px; cursor: pointer; }
.sl-outlineitem:hover { background: var(--hover); }
.sl-outlineitem.active { border-color: var(--accent); background: var(--surface); }
.sl-outlineitem .sl-sortern { position: static; }
.sl-outlinetitle { font-weight: 600; font-size: 13px; }
.sl-outlinetext { font-size: 12px; color: var(--ink-2); margin-top: 2px; white-space: pre-wrap; }
.sl-notespage { border-top: 1px solid var(--line); background: var(--surface); padding: 12px 18px; font: inherit; font-size: 13px; min-height: 160px; resize: none; outline: none; color: var(--ink); }

.sl-present { position: fixed; inset: 0; background: #000; display: grid; place-items: center; z-index: 200; }
.sl-present-stage { width: min(100vw, 177.78vh); }
.sl-present-stage svg { display: block; width: 100%; height: auto; }
.sl-present-bar {
  position: fixed; bottom: 14px; left: 50%; transform: translateX(-50%);
  color: rgba(255,255,255,0.55); font-size: 12px; letter-spacing: 0.02em;
}
`;

const SLIDE_SHORTCUTS = [
  ['Ctrl+S', 'Save'], ['Ctrl+Z', 'Undo'], ['Ctrl+N / Ctrl+O', 'New / Open'], ['Ctrl+M', 'New slide'],
  ['F5', 'Start the show from the beginning'], ['Escape', 'Leave the show'], ['→ / Space / Page Down', 'Next slide'], ['← / Page Up', 'Previous slide'],
  ['↑ / ↓', 'Previous / next slide while editing'], ['Double-click a text box', 'Edit its words'], ['Ctrl+Enter', 'Finish editing'],
];

function SlidesShortcutsDialog({ onClose }) {
  return (
    <Dialog title="Keyboard shortcuts" width={440} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      <dl className="about-list">
        {SLIDE_SHORTCUTS.map(([k, v]) => (
          <React.Fragment key={k}><dt>{k}</dt><dd>{v}</dd></React.Fragment>
        ))}
      </dl>
    </Dialog>
  );
}

/**
 * Speaker notes.
 *
 * The deck engine has always read these — an imported .pptx shows whatever its
 * author wrote — and until now there was no way to write one. They are per
 * slide, one paragraph per line, and stored in the file rather than beside it,
 * so they travel with the deck to PowerPoint and back.
 */
/** The first address a shape's words carry, or ''. */
/**
 * The heading a section draws above its first slide in the strip and the
 * sorter — its name and how many slides it holds — or nothing. A slide
 * starts a section when the one before it is in a different one, which
 * also draws a heading for a file whose sections are out of order. A click
 * goes to the slide; a double-click renames the section.
 */
function sectionHeading(model, entry, i, onRename, onGo) {
  const at = entry.section;
  if (at == null) return null;
  const before = i > 0 ? model.outline[i - 1] : null;
  if (before && before.section === at) return null;
  const section = (model.sections || [])[at];
  if (!section) return null;
  return (
    <button
      type="button"
      className="sl-section"
      data-section={at}
      title={`${section.name} — ${section.slides.length} slide${section.slides.length === 1 ? '' : 's'}. Double-click to rename.`}
      onClick={onGo}
      onDoubleClick={() => onRename(section)}
    >
      <span className="sl-section-name">{section.name}</span>
      <span className="sl-section-count">{section.slides.length}</span>
    </button>
  );
}

function linkOf(shape) {
  return (shape?.text?.paragraphs || []).flatMap((p) => p.runs || []).map((r) => (r.link && typeof r.link === 'object' ? r.link.url : null)).find(Boolean) || '';
}

/** `office.rutba.io` is an address; `office.rutba.io` with no scheme is not. */
function normaliseAddress(url) {
  const text = String(url || '').trim();
  if (/^[a-z][a-z0-9+.-]*:/i.test(text)) return text;
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(text)) return `mailto:${text}`;
  return `https://${text}`;
}

/**
 * The sections left with no slides — every slide of one deleted — drawn after
 * the last slide so they are not lost: PowerPoint keeps an empty section, and
 * so does the file. A double-click renames one; Remove All Sections takes it.
 */
function emptySectionHeadings(model, onRename) {
  return (model?.sections || []).filter((s) => !s.slides.length).map((s) => (
    <button key={`empty${s.index}`} type="button" className="sl-section" data-section={s.index} title={`${s.name} — no slides. Double-click to rename; Remove All Sections takes it away.`} onDoubleClick={() => onRename(s)}>
      <span className="sl-section-name">{s.name}</span>
      <span className="sl-section-count">0</span>
    </button>
  ));
}

/** Home → Section → Rename: the section's name, as PowerPoint asks for it. */
function SectionNameDialog({ name: current, onClose, onApply }) {
  const [name, setName] = useState(current || '');
  return (
    <Dialog
      title="Rename section"
      width={400}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Rename" className="sl-section-ok" disabled={!name.trim()} onClick={() => onApply(name.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Section name">
          <input className="rw-input sl-section-input" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && name.trim()) onApply(name.trim()); }} onFocus={(e) => e.target.select()} autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

/** Insert → Link: a web address on the selected shape's words. */
function ShapeLinkDialog({ current, words, onClose, onApply, onRemove }) {
  const [url, setUrl] = useState(current || '');
  return (
    <Dialog
      title="Link"
      width={480}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label="Remove link" className="sl-link-remove" onClick={onRemove} /> : null}
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={current ? 'Update' : 'Add link'} className="sl-link-ok" disabled={!url.trim()} onClick={() => onApply(normaliseAddress(url))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Words">
          <input className="rw-input" value={words || ''} disabled />
        </Field>
        <Field label="Address" hint="A web address, or mailto: for an email link. Ctrl+click the shape to follow it.">
          <input className="rw-input sl-link-url" value={url} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && url.trim()) onApply(normaliseAddress(url)); }} placeholder="https://office.rutba.io" autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

/** Home → Editing: the words found on every slide, each match a step to its shape, and replaced across the deck. */
function FindDialog({ onClose, onFind, onGoto, onReplaceAll }) {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [hits, setHits] = useState(null);
  const [note, setNote] = useState(null);
  const search = async () => {
    if (!find) return;
    const found = await onFind(find, matchCase);
    setHits(found);
    setNote(found.length ? null : 'Nothing matched.');
  };
  const total = (hits || []).reduce((n, h) => n + h.count, 0);
  return (
    <Dialog
      title="Find and replace"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button label="Replace all" className="sl-find-replace" disabled={!find} onClick={async () => { const found = hits ?? await onFind(find, matchCase); setNote(await onReplaceAll(find, replace, matchCase, found.reduce((n, h) => n + h.count, 0))); setHits(null); }} />
          <Button primary label="Find" className="sl-find-go" disabled={!find} onClick={search} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Find">
          <input className="rw-input sl-find-text" value={find} onChange={(e) => { setFind(e.target.value); setHits(null); }} onKeyDown={(e) => { if (e.key === 'Enter') search(); }} autoFocus />
        </Field>
        <Field label="Replace with">
          <input className="rw-input sl-find-with" value={replace} onChange={(e) => setReplace(e.target.value)} />
        </Field>
        <label className="about-auto">
          <input type="checkbox" className="sl-find-case" checked={matchCase} onChange={(e) => { setMatchCase(e.target.checked); setHits(null); }} />
          <span>Match case</span>
        </label>
        {hits?.length ? (
          <div className="ml-import-folders" style={{ maxHeight: 260 }}>
            {hits.map((h) => (
              <button key={`${h.slide}:${h.shape}`} type="button" className="ml-found-item sl-find-hit" style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }} onClick={() => onGoto(h)}>
                <span className="grow">
                  <div className="who">Slide {h.slide + 1} — {h.name}{h.count > 1 ? ` (${h.count})` : ''}</div>
                  <div className="what">{h.text}</div>
                </span>
              </button>
            ))}
          </div>
        ) : null}
        {hits?.length ? <p className="rw-hint" style={{ margin: 0 }}>{total} in {hits.length} shape{hits.length === 1 ? '' : 's'}. Click one to go to it.</p> : null}
        {note ? <div className="ml-note sl-find-note"><Icon name="info" size={14} />{note}</div> : null}
      </div>
    </Dialog>
  );
}

/** Insert → Header & Footer: the date, the slide number and the footer's words, on this slide or all of them. */
function FooterDialog({ slide, shapes, preset, onClose, onApply }) {
  const found = (type) => shapes.find((s) => s.placeholder?.type === type) || null;
  const words = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join(' ').trim();
  const dt = found('dt');
  const [dateOn, setDateOn] = useState(Boolean(dt) || preset === 'date');
  const [dateAuto, setDateAuto] = useState(!dt || Boolean(dt.text?.paragraphs?.[0]?.runs?.some((r) => r.field)));
  const [dateText, setDateText] = useState(dt && !dt.text?.paragraphs?.[0]?.runs?.some((r) => r.field) ? words(dt) : '');
  const [numberOn, setNumberOn] = useState(Boolean(found('sldNum')) || preset === 'number');
  const [footerOn, setFooterOn] = useState(Boolean(found('ftr')));
  const [footerText, setFooterText] = useState(words(found('ftr')));
  const spec = () => ({
    date: dateOn ? (dateAuto ? { auto: true } : { text: dateText }) : false,
    slideNumber: numberOn,
    footer: footerOn ? footerText : '',
  });
  const row = { display: 'flex', alignItems: 'center', gap: 8 };
  return (
    <Dialog
      title="Header and footer"
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button label="Apply to All" className="sl-hf-all" onClick={() => onApply(spec(), true)} />
          <Button primary label="Apply" className="sl-hf-apply" onClick={() => onApply(spec(), false)} />
        </>
      }
    >
      <div className="ml-form">
        <label style={row}><input type="checkbox" className="sl-hf-date-on" checked={dateOn} onChange={(e) => setDateOn(e.target.checked)} /> Date and time</label>
        <div style={{ ...row, paddingLeft: 24 }}>
          <label style={row}><input type="radio" name="sl-hf-date" className="sl-hf-date-auto" disabled={!dateOn} checked={dateAuto} onChange={() => setDateAuto(true)} /> Update automatically</label>
          <label style={row}><input type="radio" name="sl-hf-date" disabled={!dateOn} checked={!dateAuto} onChange={() => setDateAuto(false)} /> Fixed</label>
          <input className="rw-input sl-hf-date-text" style={{ flex: 1 }} disabled={!dateOn || dateAuto} value={dateText} onChange={(e) => setDateText(e.target.value)} placeholder="Spring 2026" />
        </div>
        <label style={row}><input type="checkbox" className="sl-hf-number" checked={numberOn} onChange={(e) => setNumberOn(e.target.checked)} /> Slide number</label>
        <label style={row}><input type="checkbox" className="sl-hf-footer-on" checked={footerOn} onChange={(e) => setFooterOn(e.target.checked)} /> Footer</label>
        <input className="rw-input sl-hf-footer" style={{ marginLeft: 24 }} disabled={!footerOn} value={footerText} onChange={(e) => setFooterText(e.target.value)} placeholder="The words along the bottom" autoFocus={footerOn} />
        <p className="rw-hint" style={{ margin: 0 }}>
          Slide {slide + 1} takes these with Apply; every slide with Apply to All. The number follows the slide when slides move.
        </p>
      </div>
    </Dialog>
  );
}

function NotesDialog({ slide, text, onClose, onSave }) {
  const [draft, setDraft] = useState(text || '');
  return (
    <Dialog
      title={`Speaker notes — slide ${slide + 1}`}
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Save" onClick={() => onSave(draft)} />
        </>
      }
    >
      <Field label="Only you see these" hint="They are saved into the presentation, and travel with it.">
        <textarea
          className="rw-input"
          style={{ minHeight: 200, fontFamily: 'var(--font)', resize: 'vertical' }}
          rows={9}
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          autoFocus
        />
      </Field>
    </Dialog>
  );
}