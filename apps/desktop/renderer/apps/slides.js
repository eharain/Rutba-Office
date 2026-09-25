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
import { Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Dialog, Field, Select, ZoomSlider, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { PrintDialog, defaultPrintOptions } from '../print.js';
import { usePasswordGate, openProtected, LockedAction, useProtection } from '../protect.js';
import { SITE } from '@rutba/office-formats/registry';
import Presenter, { nextShown } from './slides/presenter.js';
import SlidesRibbon from './slides/ribbon.js';
import { ShowStage, TransitionPreview, AnimationPreview } from './slides/show.js';
import { describeTransition } from './slides/motion.js';
import { clickCount } from './slides/animate.js';
import { Markup } from './slides/markup.js';
import { DesignGallery, CustomColoursDialog, CustomFontsDialog, DESIGN_CSS } from './slides/design.js';
import { CommentsPane, markerSpots, personColour, COMMENTS_CSS } from './slides/comments.js';
import { useSlidesReview } from './slides/review.js';
import { EquationDialog, EQUATION_CSS } from './word/equations.js';

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
  // A protected file whose Password dialog was cancelled: its name.
  const [lockedOut, setLockedOut] = useState(null);
  const gate = usePasswordGate();
  const [tab, setTab] = useState('home');
  /**
   * View → Slide Master: the master or layout on the stage (its part name),
   * or null in the ordinary views. Every edit made while it is set goes to
   * that part instead of a slide, and every slide on it follows.
   */
  const [masterPart, setMasterPart] = useState(null);
  const masterRef = useRef(null);
  masterRef.current = masterPart;
  /** Insert → Equation, or an equation double-clicked: { shape, initial, display } while the editor is open. */
  const [equationOpen, setEquationOpen] = useState(null);
  /** Review → the comment thread picked in the pane or on the stage, and a new one being written ({ shape }). */
  const [commentSel, setCommentSel] = useState(null);
  const [commentDraft, setCommentDraft] = useState(null);
  /** Slide Master → Rename: the part and the name it has now. */
  const [partRename, setPartRename] = useState(null);
  const [editing, setEditing] = useState(null);
  /** The shape clipboard: one shape, copied in this window, pasted on any slide of it. */
  const [clip, setClip] = useState(null);
  /** The Format Painter, armed with a shape's look: its fill, its outline and its first run's font. */
  const [painter, setPainter] = useState(null);
  const [present, setPresent] = useState(false);
  const [notesOpen, setNotesOpen] = useState(false);
  /** The Header & Footer dialog, open with a box pre-ticked ('date' | 'number') or as it stands. */
  const [footerOpen, setFooterOpen] = useState(null);
  /** Home → Find (Ctrl+F) and Replace (Ctrl+H): the pane's own mode, or closed. */
  const [findOpen, setFindOpen] = useState(false);
  /** The shape the find pane's current hit is on, so the stage can mark it apart from a plain selection. */
  const [findHit, setFindHit] = useState(null);
  /** Insert → Link on the selected shape. */
  const [linkOpen, setLinkOpen] = useState(false);
  /** The chart data dialog: the id of the chart shape it is editing, or null. */
  const [chartDataOpen, setChartDataOpen] = useState(null);
  /** Home → Section → Rename: the section (by index) and the name it has now. */
  const [sectionRename, setSectionRename] = useState(null);
  /** A shape to select once the slide a match is on has been shown. */
  const pendingSelect = useRef(null);
  const [blank, setBlank] = useState(false);
  /** Transitions → Preview (and the preview a gallery pick plays): a fresh key per run, or null. */
  const [preview, setPreview] = useState(null);
  /** Where the show is on a slide: how many of its clicks have played. */
  const [showStep, setShowStep] = useState({ index: 0, value: 0 });
  /** The show at rest — transition over, the group that was playing done — which is when After starts counting. */
  const [settled, setSettled] = useState(null);
  /** The show's stage, for finishing whatever moves when a click comes mid-animation. */
  const showControl = useRef(null);
  /** The Animations tab's current effect (its index in the slide's list), when one was picked from the pane or a badge. */
  const [animSel, setAnimSel] = useState(null);
  /** The Animation Painter, armed with a shape's effects to put on the next shape clicked. */
  const [animPainter, setAnimPainter] = useState(null);
  // Set when a presenter window is driving, so this one follows rather than leads.
  const [led, setLed] = useState(false);
  /**
   * How the deck is shown. None of it is in the file: the view mode, the
   * rulers, gridlines and guides, whether the notes strip shows, a zoom
   * level (null fits the window), and the colour/greyscale tone.
   */
  const [view, setView] = useState({ mode: 'normal', ruler: false, gridlines: false, guides: false, notes: true, zoom: null, tone: 'colour', pane: null });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  // The selection — Shift+click or Ctrl+click adds or removes a shape, a
  // click on empty stage clears it. `selectedIds[0]` is the primary — the
  // one the Font and Paragraph groups act on, and what a plain click always
  // replaces the whole selection with — so every existing read of a single
  // "selected" shape keeps working unchanged for a single selection.
  const [selectedIds, setSelectedIds] = useState([]);
  const selected = selectedIds[0] ?? null;
  const setSelected = useCallback((id) => setSelectedIds(id == null ? [] : [id]), []);
  const toggleSelected = useCallback(
    (id) => setSelectedIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id])),
    []
  );
  // A drag in progress on the stage: the shape, whether it is moved,
  // resized (and by which handle) or rotated, and the box it has been
  // dragged to, in slide units. The engine is told once, on release.
  const [drag, setDrag] = useState(null);
  // A multi-shape drag: every selected shape's id and the box it has been
  // dragged to, so the whole selection moves together as one gesture.
  const [groupDrag, setGroupDrag] = useState(null);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  /** Design → a gallery open under its button: { kind, anchor }, or null. */
  const [designOpen, setDesignOpen] = useState(null);
  /** Design → Colours → Customise Colours, and Fonts → Customise Fonts: open with the design as it stands. */
  const [customColours, setCustomColours] = useState(null);
  const [customFonts, setCustomFonts] = useState(null);
  /** The Design tab's own strips: a few themes and the four variants, drawn on this slide. */
  const [designStrip, setDesignStrip] = useState(null);
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

  /**
   * A press on a shape that is one of several selected moves the whole
   * selection together; pressing one that is not (or when only one thing is
   * selected) falls back to the single-shape drag above, which is also what
   * selects it. Released, one `setGeometry` per shape goes in a single
   * `apply` call, so the redraw and the undo step are one.
   */
  const startGroupDrag = (e, shape) => {
    if (e.button !== 0 || !shape.geometry) return;
    const ids = selectedIds.includes(shape.id) && selectedIds.length > 1 ? selectedIds : null;
    if (!ids) {
      startDrag(e, shape, 'move');
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    const g0s = new Map(ids.map((id) => [id, slide?.shapes?.find((s) => s.id === id)?.geometry]).filter(([, g]) => g).map(([id, g]) => [id, { ...g }]));
    const s = dragRef.current?.scale ?? 1;
    const x0 = e.clientX;
    const y0 = e.clientY;
    let moved = false;
    let boxes = g0s;
    const move = (ev) => {
      const dx = (ev.clientX - x0) / s;
      const dy = (ev.clientY - y0) / s;
      if (!moved && Math.abs(ev.clientX - x0) < 2 && Math.abs(ev.clientY - y0) < 2) return;
      moved = true;
      const next = new Map();
      for (const [id, g0] of g0s) next.set(id, { ...g0, x: g0.x + dx, y: g0.y + dy });
      boxes = next;
      setGroupDrag({ ids, boxes: next });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setGroupDrag(null);
      if (moved) actRef.current?.('groupDragEnd', { boxes });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  /**
   * The rotation handle above the selection: drag turns the shape about its
   * own centre, Shift snaps to 15° — the same step Alt+Left/Alt+Right nudge
   * by by from the keyboard. The handle itself orbits the shape as it turns
   * (see its own position below), so its on-screen position at the moment
   * of the press already stands for the shape's current angle; only the
   * *change* in the mouse's angle around the shape's true centre is added
   * to it, which is why the small mismatch between the pointer and the
   * handle's exact pixel does not throw the result off.
   */
  const startRotate = (e, shape) => {
    if (e.button !== 0 || !shape.geometry) return;
    e.preventDefault();
    e.stopPropagation();
    const g = shape.geometry;
    const scale = dragRef.current?.scale ?? 1;
    const offSlidePx = ROTATE_HANDLE_GAP + g.h / 2;
    const handleRect = e.currentTarget.getBoundingClientRect();
    const handleX = (handleRect.left + handleRect.right) / 2;
    const handleY = (handleRect.top + handleRect.bottom) / 2;
    const rad0 = ((g.rot || 0) * Math.PI) / 180;
    // The handle's own formula, inverted, recovers the shape's true centre
    // in screen pixels whatever the shape's current rotation already is.
    const cxScreen = handleX - offSlidePx * scale * Math.sin(rad0);
    const cyScreen = handleY + offSlidePx * scale * Math.cos(rad0);
    const startAngle = Math.atan2(e.clientY - cyScreen, e.clientX - cxScreen);
    const rot0 = g.rot || 0;
    let rot = rot0;
    const move = (ev) => {
      const angle = Math.atan2(ev.clientY - cyScreen, ev.clientX - cxScreen);
      let deg = rot0 + ((angle - startAngle) * 180) / Math.PI;
      if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
      rot = ((deg % 360) + 360) % 360;
      setDrag({ id: shape.id, kind: 'rotate', g: { ...g, rot } });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      actRef.current?.('rotateEnd', { id: shape.id, rot });
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
      // A stage taken off the page (the show replaces the editor) measures
      // nothing: keep the fit it had rather than jump to 100%.
      if (!el.isConnected || !el.clientWidth) return;
      const w = el.clientWidth - 44;
      const h = el.clientHeight - 44;
      const scale = Math.min(1, w / model.size.width, h / model.size.height);
      setFit(Number.isFinite(scale) && scale > 0 ? Math.round(scale * 1000) / 1000 : 1);
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [model?.size?.width, model?.size?.height, busy, present]);
  const menu = useMenu();
  const openFileRef = useRef(null);
  // File → Info: the Protect Presentation card and Encrypt with Password.
  const protection = useProtection({ app: 'slides', shell, doc, setDoc, toast });
  const appMenu = useAppMenu({ shell, appKey: 'slides', onNew: () => shell.win.create({ app: 'slides' }), onOpen: () => openFileRef.current?.(), extra: doc && !presenterFor ? [protection.menuItem] : [] });

  const load = useCallback(
    async (next = index) => {
      if (!doc) return;
      const m = await shell.doc.model({ id: doc.id, slide: next, width: 1280, master: masterRef.current });
      setModel(m);
    },
    [doc, index, shell, masterPart]
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
        // Slide Master view: an edit meant for "this slide" goes to the
        // master or layout on the stage; what only a slide can take is refused.
        const target = masterRef.current;
        if (target) {
          const refused = ops.find((op) => MASTER_REFUSED.has(op.op));
          if (refused) {
            toast('That works on slides — Close Master View first.', { ms: 3200 });
            return null;
          }
          ops = ops.map((op) => (typeof op.slide === 'number' ? { ...op, slide: target } : op));
        }
        const next = await shell.doc.apply({ id: doc.id, ops, slide: index, width: 1280, master: target });
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
        // A password-protected file (or the encrypted copy of one) asks
        // for its password in this window before anything opens.
        const opened = recover
          ? await openProtected((password) => shell.doc.recover({ file: recover, password }), gate)
          : boot.file
            ? await openProtected((password) => shell.doc.open({ path: boot.file, kind: 'deck', width: 1280, password }), gate)
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
        if (err.locked) setLockedOut(err.locked);
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
  }, [index, doc, masterPart]);

  // The Design tab's strips: this slide drawn in a few themes and in the
  // four variants — fetched while the tab is up, again after each change.
  useEffect(() => {
    if (tab !== 'design' || !doc?.id || !model?.count) return undefined;
    let alive = true;
    const timer = setTimeout(async () => {
      try {
        const [themes, variants] = await Promise.all([
          shell.doc.deckDesign({ id: doc.id, slide: index, kind: 'themes', width: 80 }),
          shell.doc.deckDesign({ id: doc.id, slide: index, kind: 'variants', width: 80 }),
        ]);
        if (alive) setDesignStrip({ themes: themes?.items || [], variants: variants?.items || [] });
      } catch {
        /* the strip stays as it was */
      }
    }, 120);
    return () => { alive = false; clearTimeout(timer); };
  }, [tab, doc?.id, doc?.version, index, model?.count, shell]);

  // When a presenter window is driving, this one follows: the position lives
  // in the main process precisely so the two cannot disagree about it.
  useEffect(() => {
    const off = shell.on('present:state', (s) => {
      setLed(true);
      setBlank(Boolean(s.blank));
      if (typeof s.index === 'number') setIndex(s.index);
      // The click within the slide travels with it: the presenter's Next plays the next animation here.
      setShowStep((cur) => ({ index: typeof s.index === 'number' ? s.index : cur.index, value: typeof s.step === 'number' ? s.step : 0 }));
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

  /**
   * A table cell's words, committed the same way a text box's are: each
   * line keeps the cell's own look, only the words change.
   */
  const commitTableCell = useCallback(
    async (shapeId, row, col, text) => {
      const shape = model?.slide?.shapes?.find((s) => s.id === shapeId);
      const before = shape?.table?.cells?.[row]?.[col]?.paragraphs || [];
      const paragraphs = String(text).split('\n').map((line, i) => {
        const src = before[Math.min(i, before.length - 1)] || {};
        const { runs, plain, ...props } = src;
        const look = (runs || []).find((r) => r.text && r.text !== '\n') || {};
        const { text: _text, field: _field, break: _break, link: _link, ...runProps } = look;
        return { ...props, runs: [{ ...runProps, text: line }] };
      });
      const next = await apply({ op: 'setTableCell', slide: index, shape: shapeId, row, col, paragraphs });
      setEditing(null);
      return next;
    },
    [apply, index, model]
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
    await shell.present.set({ id: doc.id, index, step: 0, running: true, blank: false });
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
      'edit.undo': { label: 'Undo', icon: 'undo', key: 'Mod+Z', run: () => actRef.current?.('undo') },
      'edit.redo': { label: 'Redo', icon: 'redo', key: 'Mod+Y', run: () => actRef.current?.('redo') },
      'edit.find': { label: 'Find…', icon: 'find', key: 'Mod+F', global: true, run: () => setFindOpen('find') },
      'edit.replace': { label: 'Replace…', icon: 'find', key: 'Mod+H', global: true, run: () => setFindOpen('replace') },
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
      // The show steps over a hidden slide rather than landing on it, and
      // through a slide's animations a click at a time before leaving it.
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown' || e.key === 'Enter') showKeys.current.next();
      if (e.key === 'ArrowLeft' || e.key === 'PageUp' || e.key === 'Backspace') showKeys.current.prev();
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

  /**
   * Where the show is: a slide and how many of its clicks have played —
   * through the presenter window's shared position when it leads, so the
   * two windows never disagree about where the show is.
   */
  const goShow = (i, value) => {
    if (led) {
      shell.present.set({ index: i, step: value }).catch(() => {});
      return;
    }
    setShowStep({ index: i, value });
    setIndex(i);
  };
  const clicksHere = () => (model?.slide?.index === index ? clickCount(model.slide.animations) : 0);
  const stepHere = () => (showStep.index === index ? Math.min(showStep.value, clicksHere()) : 0);
  /** A click, a space or an arrow: the next animation on this slide, or the next slide. A press while something moves finishes it. */
  const showNext = () => {
    if (showControl.current?.finish()) return;
    if (model?.slide?.index !== index) return;
    if (stepHere() < clicksHere()) return goShow(index, stepHere() + 1);
    const next = nextShown(model, index, 1);
    if (next !== index) goShow(next, 0);
  };
  /** Back: one click undone on this slide, or the slide before, shown as it ends. */
  const showPrev = () => {
    showControl.current?.finish();
    if (stepHere() > 0) return goShow(index, stepHere() - 1);
    const prev = nextShown(model, index, -1);
    if (prev !== index) goShow(prev, 9999);
  };
  const showKeys = useRef({ next: () => {}, prev: () => {} });
  showKeys.current = { next: showNext, prev: showPrev };
  // A show starts at the first click of its first slide.
  useEffect(() => { if (present) setShowStep({ index, value: 0 }); else setSettled(null); }, [present]); // eslint-disable-line react-hooks/exhaustive-deps

  // Advance Slide → After: once a slide is fully on screen and fully built
  // (its transition over, its animations played), the show moves on by
  // itself after that many seconds. The presenter window never runs this
  // clock; the audience window does.
  const advanceAfter = model?.slide?.transition?.advanceAfter;
  useEffect(() => {
    if (!present || presenterFor || blank || advanceAfter == null) return undefined;
    if (!settled || settled.index !== model?.slide?.index || settled.step < settled.clicks) return undefined;
    const timer = setTimeout(() => showKeys.current.next(), Math.max(0, advanceAfter * 1000));
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [present, blank, advanceAfter, settled, model?.slide?.index]);

  // A new slide means a new selection: the shape ids belong to the slide.
  useEffect(() => { setSelected(pendingSelect.current ?? null); pendingSelect.current = null; setPainter(null); setPreview(null); setAnimSel(null); setAnimPainter(null); }, [index, masterPart]);


  const slide = model?.slide;
  // Review → this slide's comment threads, and each slide's count for the strip.
  const slideComments = useMemo(() => (model?.comments || []).filter((c) => c.slide === index), [model?.comments, index]);
  const commentCount = useMemo(() => {
    const m = new Map();
    for (const c of model?.comments || []) m.set(c.slide, (m.get(c.slide) || 0) + 1);
    return m;
  }, [model?.comments]);
  const selectedShape = selected ? slide?.shapes?.find((s) => s.id === selected) || null : null;
  // Animations: the slide's sequence, and the effect the Animations tab
  // acts on — the one picked in the pane (while its shape stays
  // selected), else the selected shape's first.
  const animations = slide?.animations || [];
  const currentAnim = (() => {
    const picked = animSel != null ? animations[animSel] : null;
    if (picked && (selected == null || String(picked.shapeId) === String(selected))) return picked;
    return selected != null ? animations.find((a) => String(a.shapeId) === String(selected)) || null : null;
  })();
  const scale = view.zoom ?? fit;
  dragRef.current = { scale };
  // What the ribbon shows for the selected shape: its first run's look and
  // its first paragraph's alignment — the granularity the writer edits at.
  const format = useMemo(() => {
    const p = selectedShape?.text?.paragraphs?.[0];
    const r = p?.runs?.[0] || {};
    return {
      bold: Boolean(r.bold), italic: Boolean(r.italic), underline: Boolean(r.underline), size: r.size || selectedShape?.textDefaults?.size || 18, color: r.color || null, font: r.font || '', themeFont: selectedShape?.textDefaults?.font || null, align: p?.align || selectedShape?.textDefaults?.align || 'left',
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
        // A hidden slide never opens the show — PowerPoint starts on the first shown one.
        if (arg === 'start') setIndex(nextShown(model, -1, 1));
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
      case 'groupDragEnd': {
        // The whole selection, dragged together: one setGeometry per shape,
        // all in the same apply — one redraw, one undo step.
        const ops = [...arg.boxes.entries()].map(([id, g]) => ({ op: 'setGeometry', slide: index, shape: id, x: Math.round(g.x), y: Math.round(g.y), w: Math.round(g.w), h: Math.round(g.h) }));
        if (ops.length) await apply(...ops);
        return;
      }
      case 'rotateEnd':
        await apply({ op: 'rotateShapes', slide: index, ids: [arg.id], delta: arg.rot - (slide?.shapes?.find((s) => s.id === arg.id)?.geometry?.rot || 0) });
        return;
      // Home → Arrange → Align/Distribute/Rotate/Flip/Group/Ungroup — every
      // one an op on the whole selection (a group counts as one shape).
      case 'align': {
        const ids = selectedIds.length ? selectedIds : (selectedShape ? [selectedShape.id] : []);
        if (!ids.length) return toast('Select a shape first.', { ms: 3000 });
        await apply({ op: 'alignShapes', slide: index, ids, edge: arg.edge, to: arg.to });
        return;
      }
      case 'distribute': {
        if (selectedIds.length < 3) return toast('Select three or more shapes first.', { ms: 3000 });
        await apply({ op: 'distributeShapes', slide: index, ids: selectedIds, axis: arg.axis, to: arg.to });
        return;
      }
      case 'rotateBy': {
        const ids = selectedIds.length ? selectedIds : (selectedShape ? [selectedShape.id] : []);
        if (!ids.length) return;
        await apply({ op: 'rotateShapes', slide: index, ids, delta: arg });
        return;
      }
      case 'flipShape': {
        const ids = selectedIds.length ? selectedIds : (selectedShape ? [selectedShape.id] : []);
        if (!ids.length) return;
        await apply({ op: 'flipShapes', slide: index, ids, axis: arg });
        return;
      }
      case 'group': {
        if (selectedIds.length < 2) return toast('Select two or more shapes to group.', { ms: 3000 });
        const next = await apply({ op: 'groupShapes', slide: index, ids: selectedIds });
        // Shape ids are strings on the scene; `groupShapes` hands back the
        // number it minted the id from, so the newly-drawn group's own
        // (string) id is looked up rather than trusted to match `===`.
        const gid = next?.model?.slide?.shapes?.find((s) => String(s.id) === String(next?.opResult))?.id;
        if (gid != null) setSelected(gid);
        return;
      }
      case 'ungroup': {
        if (!selectedShape || selectedShape.kind !== 'group') return toast('Select a group to ungroup.', { ms: 3000 });
        const next = await apply({ op: 'ungroupShape', slide: index, shape: selectedShape.id });
        const members = typeof next?.opResult === 'string' ? next.opResult.split(',').filter(Boolean) : [];
        if (members.length) setSelectedIds(members);
        return;
      }
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
        // Merged with whatever effects the shape already has — a shadow
        // preset turns the shadow on or off without disturbing a glow, a
        // reflection or soft edges the shape also carries, the same way
        // 'shapeEffects' (below) merges the others in. The shape's current
        // effects come back from the scene in its own read units, so they
        // go through effectsToWriteSpec first — merging them in as read
        // silently drops whatever setShapeStyle's writer names differently.
        if (!selectedShape || !(arg in SHADOW_PRESETS)) return;
        const effects = { ...effectsToWriteSpec(selectedShape.effects), shadow: SHADOW_PRESETS[arg] };
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, effects });
        return;
      }
      // Shape Effects → Glow / Soft Edges / Reflection: `arg` names the one
      // key to change (`{ glow: {...} | null }`, and so on); everything
      // else the shape's effects already carry rides along unchanged, via
      // effectsToWriteSpec — see its own comment.
      case 'shapeEffects': {
        if (!selectedShape) return;
        const effects = { ...effectsToWriteSpec(selectedShape.effects), ...arg };
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, effects });
        return;
      }
      // The Format pane's Picture fill: a picture from this device, the
      // same dialog Insert → Pictures uses, filling the selected shape and
      // clipped to its outline.
      case 'pickPictureFill': {
        if (!selectedShape) return;
        const [file] = await shell.dialog.open({
          title: 'Choose a picture',
          filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp'] }],
        });
        if (!file) return;
        const { bytes, stat } = await shell.fs.read({ path: file });
        const ext = String(stat?.ext || file.split('.').pop()).replace('.', '').toLowerCase();
        const contentType = { png: 'image/png', gif: 'image/gif', bmp: 'image/bmp' }[ext] || 'image/jpeg';
        await apply({ op: 'setShapeStyle', slide: index, shape: selectedShape.id, fill: { picture: { data: bytes, contentType, name: stat?.name || file.split(/[\/]/).pop() } } });
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
      // Slide Show → Hide Slide: left out of the show; everything else still draws it.
      case 'hideSlide':
        await apply({ op: 'setSlideHidden', slide: index, hidden: !model?.slide?.hidden });
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
      // Design → Background Styles: this slide's own background, or every slide's.
      case 'background':
        await apply({ op: 'setBackground', slide: index, spec: arg.spec ?? null, all: Boolean(arg.all) });
        return;
      // Design → Themes, Variants, Colours, Fonts, Effects: a gallery under
      // its button, and a pick from it — one op on the deck, one undo step.
      case 'designGallery': setDesignOpen(arg); return;
      case 'applyTheme': {
        setDesignOpen(null);
        const next = await apply({ op: 'applyTheme', theme: arg.id ?? arg, variant: arg.variant ?? 0 });
        if (next) toast(`${arg.name || 'Theme'} applied to every slide`, { ms: 2200 });
        return;
      }
      case 'applyVariant':
        setDesignOpen(null);
        await apply({ op: 'applyVariant', variant: arg, slide: index });
        return;
      case 'themeColors':
        setDesignOpen(null);
        await apply({ op: 'setThemeColors', palette: arg });
        return;
      case 'themeFonts':
        setDesignOpen(null);
        await apply({ op: 'setThemeFonts', pair: arg });
        return;
      case 'themeEffects':
        setDesignOpen(null);
        await apply({ op: 'setThemeEffects', effects: arg });
        return;
      case 'customColours': setDesignOpen(null); setCustomColours(arg || model?.design || {}); return;
      case 'customFonts': setDesignOpen(null); setCustomFonts(arg || model?.design || {}); return;
      // Insert → Equation: Word's equation editor, over the slide.
      case 'equation':
        if (masterPart) return toast('Equations go on slides — Close Master View first.', { ms: 3200 });
        setEquationOpen({ shape: null, initial: '', display: true });
        return;
      case 'putEquation': {
        const open = equationOpen || {};
        try {
          const next = await shell.doc.deckEquation({ id: doc.id, slide: index, shape: open.shape ?? null, linear: arg.linear, display: arg.display, width: 1280 });
          setDoc(next);
          setModel(next.model);
          setEquationOpen(null);
          if (next.opResult != null) setSelected(String(next.opResult));
        } catch (err) {
          toast(err.message, { tone: 'bad' });
        }
        return;
      }
      // Review → Comments: a new thread, a post, a reply, Resolve, Delete,
      // Previous and Next across the deck, and the pane.
      case 'newComment':
        if (masterPart) return toast('Comments are on slides — Close Master View first.', { ms: 3200 });
        setCommentDraft({ shape: selected ?? null });
        setCommentSel(null);
        patchView({ pane: 'comments' });
        return;
      case 'postComment': {
        const draft = commentDraft || { shape: null };
        const next = await apply({ op: 'addComment', slide: index, text: arg, shape: draft.shape ?? null });
        if (next) {
          setCommentDraft(null);
          if (typeof next.opResult === 'string') setCommentSel(next.opResult);
        }
        return;
      }
      case 'replyComment':
        await apply({ op: 'replyComment', slide: arg.thread.slide, id: arg.thread.id, text: arg.text });
        return;
      case 'resolveComment':
        await apply({ op: 'resolveComment', slide: arg.thread.slide, id: arg.thread.id, resolved: arg.resolved });
        return;
      case 'deleteComment': {
        const t = arg || (model?.comments || []).find((c) => c.id === commentSel);
        if (!t) return toast('Pick a comment first — in the pane or on the slide.', { ms: 3000 });
        const next = await apply({ op: 'removeComment', slide: t.slide, id: t.id });
        if (next && commentSel === t.id) setCommentSel(null);
        return;
      }
      case 'deleteReply':
        await apply({ op: 'removeComment', slide: arg.thread.slide, id: arg.thread.id, reply: arg.reply.id });
        return;
      case 'deleteComments': {
        const next = await apply({ op: 'removeAllComments', slide: index, all: arg === 'all' });
        if (next) { setCommentSel(null); toast(arg === 'all' ? 'Every comment in the presentation deleted' : 'Every comment on this slide deleted', { ms: 2400 }); }
        return;
      }
      case 'commentStep': {
        const list = model?.comments || [];
        if (!list.length) return toast('There are no comments in this presentation.', { ms: 2600 });
        const at = list.findIndex((c) => c.id === commentSel);
        const target = at >= 0
          ? list[(at + arg + list.length) % list.length]
          : arg > 0 ? list.find((c) => c.slide >= index) || list[0] : [...list].reverse().find((c) => c.slide <= index) || list[list.length - 1];
        setCommentSel(target.id);
        setCommentDraft(null);
        if (target.slide !== index) setIndex(target.slide);
        patchView({ pane: 'comments' });
        return;
      }
      case 'selectComment':
        setCommentSel(arg);
        patchView({ pane: 'comments' });
        return;
      // View → Slide Master, and the Slide Master tab's verbs.
      case 'masterView': {
        const layout = model?.slide?.layout || null;
        setSelected(null);
        setEditing(null);
        setMasterPart(layout || model?.slide?.master || 'ppt/slideMasters/slideMaster1.xml');
        setTab('master');
        return;
      }
      case 'closeMaster':
        setMasterPart(null);
        setSelected(null);
        setEditing(null);
        setTab('home');
        return;
      case 'masterSelect': setMasterPart(arg); return;
      case 'insertLayout': {
        const master = model?.masterView?.items?.find((it) => it.part === masterPart)?.master || model?.slide?.master || null;
        const next = await apply({ op: 'insertLayout', master, name: 'Custom Layout' });
        if (next && typeof next.opResult === 'string') setMasterPart(next.opResult);
        return;
      }
      case 'renameLayout': {
        const item = model?.masterView?.items?.find((it) => it.part === masterPart);
        if (item) setPartRename({ part: item.part, name: item.name, kind: item.kind });
        return;
      }
      case 'deleteLayout': {
        const item = model?.masterView?.items?.find((it) => it.part === masterPart);
        if (!item || item.kind !== 'layout') return;
        const next = await apply({ op: 'removeLayout', part: item.part });
        if (next) { setMasterPart(item.master); toast(`Layout "${item.name}" deleted`, { ms: 2200 }); }
        return;
      }
      case 'masterPlaceholders':
        await apply({ op: 'setMasterPlaceholders', part: masterPart, ...arg });
        return;
      case 'hideBackgroundGraphics':
        await apply({ op: 'setHideBackgroundGraphics', part: masterPart, hide: Boolean(arg) });
        return;
      case 'preserve': {
        const item = model?.masterView?.items?.find((it) => it.part === masterPart);
        const master = item?.kind === 'master' ? item.part : item?.master;
        const on = !model?.masterView?.items?.find((it) => it.part === master)?.preserve;
        if (master) await apply({ op: 'setMasterPreserve', part: master, on });
        return;
      }
      case 'insertPlaceholder': {
        const next = await apply({ op: 'insertPlaceholder', part: masterPart, kind: arg });
        if (next?.opResult != null) setSelected(String(next.opResult));
        return;
      }
      case 'undo': {
        if (!doc?.canUndo) return;
        const n = await shell.doc.undo({ id: doc.id, slide: index, width: 1280, master: masterPart });
        setDoc(n);
        setModel(n.model);
        return;
      }
      case 'redo': {
        if (!doc?.canRedo) return;
        const n = await shell.doc.redo({ id: doc.id, slide: index, width: 1280, master: masterPart });
        setDoc(n);
        setModel(n.model);
        return;
      }
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
      case 'addTable': {
        // Centred, PowerPoint's own default size and style; selected, so
        // Arrange and Delete act on it at once.
        const next = await apply({ op: 'addTable', slide: index, rows: arg.rows, cols: arg.cols });
        const added = next?.model?.slide?.shapes?.slice(-1)[0];
        if (added) setSelected(added.id);
        return;
      }
      case 'addChart': {
        // A sample chart — Q1..Q4, Sales and Costs — centred at the engine's
        // own default size; selected, so Arrange and Delete act on it at
        // once. A pie takes only the first series: two slices of the same
        // pie are one measure, not two.
        const categories = ['Q1', 'Q2', 'Q3', 'Q4'];
        const series = arg.type === 'pie'
          ? [{ name: 'Sales', values: [12, 18, 15, 22] }]
          : [{ name: 'Sales', values: [12, 18, 15, 22] }, { name: 'Costs', values: [8, 9, 10, 11] }];
        const next = await apply({ op: 'addChart', slide: index, type: arg.type, categories, series });
        const added = next?.model?.slide?.shapes?.slice(-1)[0];
        if (added) setSelected(added.id);
        return;
      }
      // Right-click a cell: a row or column added or taken away.
      case 'tableRow':
        await apply({ op: arg.remove ? 'removeTableRow' : 'insertTableRow', slide: index, shape: arg.shape, at: arg.at });
        return;
      case 'tableColumn':
        await apply({ op: arg.remove ? 'removeTableColumn' : 'insertTableColumn', slide: index, shape: arg.shape, at: arg.at });
        return;
      case 'footer': {
        setFooterOpen(arg || 'open');
        return;
      }
      case 'find': {
        setFindOpen('find');
        return;
      }
      case 'replace': {
        setFindOpen('replace');
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
        if (!selectedIds.length) return;
        // Every selected shape (a group among them counts as one), in one
        // apply, so the deletion is a single undo step.
        await apply(...selectedIds.map((id) => ({ op: 'removeShape', slide: index, shape: id })));
        setSelected(null);
        return;
      case 'nudge': {
        const ids = selectedIds.length ? selectedIds : (selectedShape ? [selectedShape.id] : []);
        const ops = ids
          .map((id) => slide?.shapes?.find((s) => s.id === id))
          .filter((s) => s?.geometry)
          .map((s) => ({ op: 'setGeometry', slide: index, shape: s.id, x: s.geometry.x + (arg.dx || 0), y: s.geometry.y + (arg.dy || 0), w: s.geometry.w, h: s.geometry.h }));
        if (ops.length) await apply(...ops);
        return;
      }
      case 'format': {
        if (!selectedShape?.text) return toast('Click a text box first.', { ms: 3500 });
        // Slide Master view: a placeholder's look is a text style — the
        // master's title or body style, or the placeholder's own list style —
        // which every slide on it inherits; its prompt words are not touched.
        if (masterPart && selectedShape.placeholder) {
          const d = selectedShape.textDefaults || {};
          const props = {};
          if ('size' in arg && arg.size != null) props.size = arg.size;
          if ('color' in arg && arg.color) props.color = arg.color;
          if ('font' in arg) props.font = arg.font || null;
          if ('align' in arg && arg.align) props.align = arg.align;
          if ('bold' in arg) props.bold = arg.bold === 'toggle' ? !d.bold : Boolean(arg.bold);
          if ('italic' in arg) props.italic = arg.italic === 'toggle' ? !d.italic : Boolean(arg.italic);
          if ('underline' in arg) props.underline = arg.underline === 'toggle' ? !d.underline : Boolean(arg.underline);
          if (Object.keys(props).length) {
            await apply({ op: 'setTextStyle', slide: index, shape: selectedShape.id, props });
            return;
          }
          return toast('In Slide Master view the font, size, colour, weight and alignment of a placeholder are its text style; the rest belongs to a slide.', { ms: 4200 });
        }
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
      // Transitions → the gallery, Effect Options, Duration and Advance
      // Slide: one op merging what changed over this slide's transition. A
      // pick from the gallery or its options plays itself on the stage, the
      // way PowerPoint previews an effect as it is chosen.
      case 'transition': {
        const next = await apply({ op: 'setTransition', slide: index, spec: arg });
        if (next && (arg.type !== undefined || arg.direction !== undefined) && arg.type !== 'none') setPreview({ kind: 'transition', key: Date.now() });
        return;
      }
      case 'transitionAll': {
        const next = await apply({ op: 'applyTransitionToAll', slide: index });
        if (next) toast(next.opResult ? `This slide's transition is now on every slide` : 'Every slide already has this transition', { tone: 'good' });
        return;
      }
      case 'preview':
        setPreview({ kind: arg || 'transition', key: Date.now(), only: null });
        return;
      // Animations → the gallery: the selected shape's current effect
      // becomes this one (or, for a shape with none, gains it); None takes
      // every effect off the selected shapes. Each selected shape gets the
      // same, in one apply; the effect made previews itself on the stage.
      case 'animate':
      case 'addAnimation': {
        const ids = selectedIds.length ? selectedIds : selected ? [selected] : [];
        if (!ids.length) return toast('Click a shape first.', { ms: 3500 });
        if (arg.effect === 'none') {
          await apply(...ids.map((id) => ({ op: 'removeShapeAnimations', slide: index, shape: id })));
          setAnimSel(null);
          return;
        }
        const ops = ids.map((id) => {
          const own = name === 'animate' ? (String(id) === String(selected) && currentAnim ? currentAnim : animations.find((a) => String(a.shapeId) === String(id))) : null;
          return own
            ? { op: 'setAnimation', slide: index, index: own.index, patch: { kind: arg.kind, effect: arg.effect } }
            : { op: 'addAnimation', slide: index, shape: id, spec: { kind: arg.kind, effect: arg.effect } };
        });
        const next = await apply(...ops);
        if (!next) return;
        const last = ops[ops.length - 1];
        const at = last.op === 'addAnimation' ? next.opResult : last.index;
        if (typeof at === 'number') {
          setAnimSel(at);
          setPreview({ kind: 'animation', key: Date.now(), only: at });
        }
        return;
      }
      // Effect Options, Start, Duration and Delay: the current effect changed.
      case 'animPatch': {
        if (!currentAnim) return;
        await apply({ op: 'setAnimation', slide: index, index: currentAnim.index, patch: arg });
        if (arg.direction !== undefined) setPreview({ kind: 'animation', key: Date.now(), only: currentAnim.index });
        return;
      }
      // Move Earlier / Move Later, or a row dropped elsewhere in the pane.
      case 'animMove': {
        const from = arg?.from ?? currentAnim?.index;
        if (from == null) return;
        const next = await apply({ op: 'moveAnimation', slide: index, index: from, to: arg?.to ?? arg });
        if (typeof next?.opResult === 'number') setAnimSel(next.opResult);
        return;
      }
      case 'animRemove': {
        const at = arg ?? currentAnim?.index;
        if (at == null) return;
        await apply({ op: 'removeAnimation', slide: index, index: at });
        setAnimSel(null);
        return;
      }
      // A row in the Animation Pane or a number on the stage: that effect, and its shape selected.
      case 'animSelect': {
        const e = animations[arg];
        if (!e) return;
        setAnimSel(arg);
        if (e.shapeId != null) setSelected(e.shapeId);
        return;
      }
      // A row's own menu in the Animation Pane: its start, or out.
      case 'animMenu': {
        const e = animations[arg.index];
        if (!e) return;
        menu.open(arg.ev, [
          ...[['onClick', 'Start On Click'], ['withPrevious', 'Start With Previous'], ['afterPrevious', 'Start After Previous']].map(([trigger, label]) => ({
            label,
            icon: e.trigger === trigger ? 'check' : undefined,
            run: () => apply({ op: 'setAnimation', slide: index, index: arg.index, patch: { trigger } }),
          })),
          '-',
          { label: 'Remove', icon: 'trash', run: () => act('animRemove', arg.index) },
        ]);
        return;
      }
      // The Animation Painter: armed with the selected shape's effects, put
      // on the next shape clicked in place of its own.
      case 'animPainter': {
        if (animPainter) { setAnimPainter(null); return; }
        const own = animations.filter((a) => String(a.shapeId) === String(selected) && a.known);
        if (!own.length) return toast('That shape has no animation to copy.', { ms: 3500 });
        setAnimPainter(own.map(({ kind, effect, direction, trigger, duration, delay }) => ({ kind, effect, direction, trigger, duration, delay })));
        return;
      }
      case 'animPaint': {
        const specs = animPainter;
        setAnimPainter(null);
        if (!specs || !arg) return;
        await apply(
          { op: 'removeShapeAnimations', slide: index, shape: arg.id },
          ...specs.map((spec) => ({ op: 'addAnimation', slide: index, shape: arg.id, spec })),
        );
        setSelected(arg.id);
        setAnimSel(null);
        return;
      }
      default:
        toast(`${name} is not wired yet.`, { ms: 3000 });
    }
  };

  actRef.current = act;
  // Review → Check Accessibility and Spelling (slides/review.js).
  const review = useSlidesReview({ shell, doc, model, apply, toast, index, setIndex, setSelected, patchView });

  // Every hook above, every early return below. This return sat above the
  // formatting memo, so a file the engine refused made React throw
  // "rendered fewer hooks than expected" and the person got a blank window
  // instead of the reason — for a truncated deck, twenty seconds of nothing.
  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Presentation" menu={appMenu}>
        <Empty icon={lockedOut ? 'lock' : 'slides'} title={lockedOut ? 'This presentation is password-protected' : 'This file could not be opened'} action={lockedOut ? <LockedAction /> : null}>{error}</Empty>
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
        onClick={() => {
          // Advance Slide → On Mouse Click off: a click still plays this
          // slide's animations but does not move the show off it (the keys
          // still do), as in PowerPoint.
          if (slide.transition?.advanceOnClick === false && stepHere() >= clicksHere()) return;
          showNext();
        }}
      >
        <style>{CSS}</style>
        {/* A black screen is a thing speakers ask for by name: attention back on them. It hides the stage rather than dropping it, so coming back does not replay the transition. */}
        <ShowStage slide={slide} size={model.size} step={showStep} hidden={blank} onSettled={setSettled} control={showControl} />
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
          selectedIds={selectedIds}
          format={format}
          canPaste={Boolean(clip)}
          painter={Boolean(painter)}
          animation={currentAnim}
          animPainter={Boolean(animPainter)}
          addSlide={addSlide}
          insertPicture={insertPicture}
          presentWithNotes={presentWithNotes}

          setPresent={setPresent}
          setNotesOpen={setNotesOpen}
          designStrip={designStrip}
          masterView={model?.masterView || null}
          masterPart={masterPart}
          review={review}
        />
      }
      status={
        <>
          <span>{doc?.path || 'Not saved yet'}</span>
          <Spacer />
          {model?.masterView ? (() => {
            const item = model.masterView.items.find((it) => it.part === masterPart);
            return <Chip>{item?.kind === 'layout' ? `${item.name} layout: used by ${item.used === 1 ? '1 slide' : `${item.used} slides`}` : `${item?.name || 'Slide Master'}: used by every layout`}</Chip>;
          })() : <Chip>Slide {index + 1} of {model?.count ?? 0}</Chip>}
          {slide?.shapes ? <Chip>{slide.shapes.length} shapes</Chip> : null}
          {review.status}
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
          <style>{CSS + DESIGN_CSS + COMMENTS_CSS + EQUATION_CSS}</style>
          <Panel width={196} resizable title={model.masterView ? 'Slide Master' : 'Slides'}>
            {model.masterView ? (
              <div className="sl-sorter sl-masterstrip">
                {model.masterView.items.map((it) => (
                  <button
                    key={it.part}
                    type="button"
                    className={`sl-thumb sl-mthumb sl-mthumb-${it.kind}${it.part === masterPart ? ' active' : ''}`}
                    data-part={it.part}
                    onClick={() => act('masterSelect', it.part)}
                    onContextMenu={(e) => {
                      act('masterSelect', it.part);
                      menu.open(e, [
                        { label: 'Insert Layout', icon: 'plus', run: () => act('insertLayout') },
                        { label: 'Rename…', icon: 'textbox', run: () => setPartRename({ part: it.part, name: it.name, kind: it.kind }) },
                        ...(it.kind === 'layout' ? [{ label: it.used ? `Delete (used by ${it.used === 1 ? '1 slide' : `${it.used} slides`})` : 'Delete Layout', icon: 'trash', disabled: it.used > 0, run: () => act('deleteLayout') }] : []),
                      ]);
                    }}
                  >
                    <span className="sl-thumb-card" data-tip={it.kind === 'master' ? `${it.name} — every layout below takes its look from here` : `${it.name} Layout: used by ${it.used === 1 ? '1 slide' : `${it.used} slides`}`}>
                      {it.thumbnail ? <Markup as="span" className="sl-thumb-pic" html={it.thumbnail} /> : <span className="sl-thumb-title">{it.name}</span>}
                    </span>
                    <span className="sl-mthumb-name">{it.name}{it.kind === 'layout' && it.used ? <span className="sl-mthumb-used">{it.used}</span> : null}</span>
                  </button>
                ))}
              </div>
            ) : (
            <div className="sl-sorter">
              {(model.outline || []).map((o, i) => (
                <React.Fragment key={o.part || i}>
                  {sectionHeading(model, o, i, (s) => setSectionRename({ section: s.index, name: s.name }), () => setIndex(i))}
                  <button
                    type="button"
                    className={`sl-thumb${i === index ? ' active' : ''}${o.hidden ? ' hidden' : ''}`}
                    onClick={() => setIndex(i)}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['slide.new', 'slide.delete']))}
                  >
                    <span className="sl-thumb-n">
                      {i + 1}
                      {/* PowerPoint's little star under the number: this slide has a transition. */}
                      {commentCount.get(i) ? <span className="sl-thumb-cm" data-comments={commentCount.get(i)} title={`${commentCount.get(i)} comment${commentCount.get(i) === 1 ? '' : 's'}`}><Icon name="reply" size={10} /></span> : null}
                      {o.transition || o.animated ? <span className="sl-thumb-fx" data-fx={[o.transition ? 'transition' : null, o.animated ? 'animations' : null].filter(Boolean).join(' ')} title={[o.transition ? `Transition: ${describeTransition({ type: o.transition })}` : null, o.animated ? 'Has animations' : null].filter(Boolean).join(' · ')}><Icon name="star" size={10} /></span> : null}
                    </span>
                    <span className="sl-thumb-card" title={`${o.title || `Slide ${i + 1}`}${o.hidden ? ' — hidden' : ''}`}>
                      {o.thumbnail
                        ? <Markup as="span" className="sl-thumb-pic" html={o.thumbnail} />
                        : <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>}
                    </span>
                  </button>
                </React.Fragment>
              ))}
              {emptySectionHeadings(model, (s) => setSectionRename({ section: s.index, name: s.name }))}
            </div>
            )}
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
                // Alt+Left/Alt+Right: PowerPoint's own keyboard rotate, 15° a press.
                if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
                  e.preventDefault();
                  act('rotateBy', e.key === 'ArrowLeft' ? -15 : 15);
                  return;
                }
                const step = e.shiftKey ? 10 : 1;
                const nudge = { ArrowLeft: { dx: -step }, ArrowRight: { dx: step }, ArrowUp: { dy: -step }, ArrowDown: { dy: step } }[e.key];
                if (nudge) { e.preventDefault(); act('nudge', nudge); }
                else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); act('deleteShape'); }
                else if (e.key === 'Escape') { if (painter) setPainter(null); else if (animPainter) setAnimPainter(null); else setSelected(null); }
              }}
            >
              {findOpen ? (
                <FindPane
                  key={findOpen}
                  mode={findOpen}
                  onClose={() => { setFindOpen(false); setFindHit(null); }}
                  onSearch={(text, matchCase) => (doc ? shell.doc.deckFind({ id: doc.id, query: text, options: { matchCase } }) : Promise.resolve([]))}
                  onGoto={(hit) => {
                    setFindHit(hit ? hit.shape : null);
                    if (!hit) return;
                    if (hit.slide === index) setSelected(hit.shape);
                    else { pendingSelect.current = hit.shape; setIndex(hit.slide); }
                  }}
                  onReplaceOne={async (hit, replacement) => {
                    const next = await apply({ op: 'replaceHit', hit, replacement });
                    return Boolean(next?.opResult);
                  }}
                  onReplaceAll={async (find, replace, matchCase) => {
                    const next = await apply({ op: 'replaceAllHits', find, replace, matchCase });
                    return next?.opResult ?? 0;
                  }}
                />
              ) : null}
              {slide && !masterPart && view.mode === 'sorter' ? (
                <div className="sl-sortergrid">
                  {(model.outline || []).map((o, i) => (
                    <React.Fragment key={o.part || i}>
                      {sectionHeading(model, o, i, (s) => setSectionRename({ section: s.index, name: s.name }), () => setIndex(i))}
                      <button type="button" className={`sl-sortercard${i === index ? ' active' : ''}${o.hidden ? ' hidden' : ''}`} onClick={() => { setIndex(i); patchView({ mode: 'normal' }); }} title={`${o.title || `Slide ${i + 1}`}${o.hidden ? ' — hidden' : ''}`}>
                        {o.thumbnail ? <Markup as="span" className="sl-thumb-pic" html={o.thumbnail} /> : <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>}
                        <span className="sl-sortern">{i + 1}</span>
                        {o.transition || o.animated ? <span className="sl-sorterfx" title={o.transition && o.animated ? 'This slide has a transition and animations' : o.transition ? 'This slide has a transition' : 'This slide has animations'}><Icon name="star" size={11} /></span> : null}
                      </button>
                    </React.Fragment>
                  ))}
                  {emptySectionHeadings(model, (s) => setSectionRename({ section: s.index, name: s.name }))}
                </div>
              ) : slide && !masterPart && view.mode === 'outline' ? (
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
                  <Markup className="sl-svg" html={slide.svg} />
                  {/* Transitions → Preview: the slide before (or black) into this one, over the stage. */}
                  {preview?.kind === 'transition' && slide.transition ? (
                    <TransitionPreview key={preview.key} fromSvg={index > 0 ? model.outline?.[index - 1]?.thumbnail || '' : ''} toSvg={slide.svg} transition={slide.transition} onDone={() => setPreview(null)} />
                  ) : null}
                  {/* Animations → Preview: the slide's sequence (or the effect just picked) played over the stage. */}
                  {preview?.kind === 'animation' && animations.length ? (
                    <AnimationPreview key={preview.key} slide={slide} size={model.size} only={preview.only} onDone={() => setPreview(null)} />
                  ) : null}
                  {view.gridlines ? <div className="sl-gridlines" /> : null}
                  {view.guides ? <div className="sl-guides" /> : null}
                  {/*
                    Every top-level shape gets a hit area over the drawing —
                    a shape inside a group does not get one of its own, so a
                    click anywhere on a group's members selects the group,
                    the way PowerPoint's own does. Click selects (Shift or
                    Ctrl adds or removes one from the selection), drag moves
                    (the whole selection together, when the one pressed is
                    already part of it), a double-click on words edits them.
                    A single selection wears eight handles and a rotation
                    handle; several share one dashed frame.
                  */}
                  {slide.shapes.filter((s) => s.geometry && !s.hidden && s.groupId == null).map((s) => {
                    const g = drag?.id === s.id ? drag.g : groupDrag?.boxes?.get(s.id) ?? s.geometry;
                    return (
                      <button
                        key={s.id}
                        type="button"
                        className={`sl-hit${selectedIds.includes(s.id) ? ' selected' : ''}${selected === s.id ? ' primary' : ''}${drag?.id === s.id || groupDrag?.ids?.includes(s.id) ? ' dragging' : ''}${findHit === s.id ? ' find-current' : ''}`}
                        data-shape={s.id}
                        style={{ left: g.x, top: g.y, width: g.w, height: g.h }}
                        onMouseDown={(e) => { if (e.shiftKey || e.ctrlKey || e.metaKey) return; startGroupDrag(e, s); }}
                        onClick={(e) => {
                          if ((e.ctrlKey || e.metaKey) && linkOf(s)) { shell.shell.openExternal({ url: linkOf(s) }); return; }
                          if (painter) { paintShape(s); return; }
                          if (animPainter) { act('animPaint', s); return; }
                          if (e.shiftKey || e.ctrlKey || e.metaKey) { toggleSelected(s.id); return; }
                          setSelected(s.id);
                        }}
                        onDoubleClick={() => {
                          if (s.kind === 'chart') { setSelected(s.id); setChartDataOpen(s.id); return; }
                          // An equation opens in the equation editor, in its linear form.
                          const eq = (s.text?.paragraphs || []).flatMap((p) => p.runs || []).find((r) => r.math);
                          if (eq && !masterPart) { setSelected(s.id); setEquationOpen({ shape: s.id, initial: eq.math.linear || '', display: eq.math.display !== false }); return; }
                          if (s.text) setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') });
                        }}
                        onContextMenu={(e) => {
                          if (!selectedIds.includes(s.id)) setSelected(s.id);
                          menu.open(e, [
                            ...(s.text ? [{ label: 'Edit text', icon: 'textbox', run: () => setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') }) }] : []),
                            ...(s.kind === 'chart' ? [{ label: 'Edit Data…', icon: 'table', run: () => { setSelected(s.id); setChartDataOpen(s.id); } }] : []),
                            { label: 'Format shape…', icon: 'wand', run: () => { setSelected(s.id); act('formatPane'); } },
                            { label: 'Edit Alt Text…', icon: 'textbox', run: () => { setSelected(s.id); review.openAltText({ slide: index, shape: s.id }); } },
                            { label: 'Bring to front', icon: 'chevronUp', run: () => { setSelected(s.id); apply({ op: 'reorderShape', slide: index, shape: s.id, to: 'front' }); } },
                            { label: 'Send to back', icon: 'chevronDown', run: () => { setSelected(s.id); apply({ op: 'reorderShape', slide: index, shape: s.id, to: 'back' }); } },
                            '-',
                            { label: 'Group', icon: 'grid', disabled: selectedIds.length < 2, run: () => act('group') },
                            { label: 'Ungroup', icon: 'grid', disabled: s.kind !== 'group', run: () => act('ungroup') },
                            '-',
                            { label: 'Delete shape', icon: 'trash', run: () => act('deleteShape') },
                          ]);
                        }}
                        title={(s.text ? `${s.name || 'Shape'} — drag to move, double-click to edit` : `${s.name || s.kind} — drag to move`) + (linkOf(s) ? ` — Ctrl+click to follow ${linkOf(s)}` : '')}
                      />
                    );
                  })}
                  {/*
                    A table draws its own grid in the SVG; a transparent hit
                    layer per cell sits over it, so a double-click edits that
                    cell in place and a right-click offers the row and
                    column verbs, the way PowerPoint's own table does.
                  */}
                  {slide.shapes.filter((s) => s.kind === 'table' && s.table && !s.hidden).flatMap((s) =>
                    s.table.cells.flatMap((row, ri) => row.map((cell, ci) => (
                      <div
                        key={`${s.id}-${ri}-${ci}`}
                        className="sl-cell-hit"
                        data-row={ri}
                        data-col={ci}
                        style={cell.box ? { left: cell.box.x, top: cell.box.y, width: cell.box.w, height: cell.box.h } : {}}
                        onMouseDown={(e) => { setSelected(s.id); startDrag(e, s, 'move'); }}
                        onDoubleClick={(e) => {
                          e.stopPropagation();
                          setSelected(s.id);
                          setEditing({ id: s.id, row: ri, col: ci, text: (cell.paragraphs || []).map((p) => p.plain).join('\n') });
                        }}
                        onContextMenu={(e) => {
                          setSelected(s.id);
                          menu.open(e, [
                            { label: 'Insert row above', run: () => act('tableRow', { shape: s.id, at: ri }) },
                            { label: 'Insert row below', run: () => act('tableRow', { shape: s.id, at: ri + 1 }) },
                            { label: 'Insert column left', run: () => act('tableColumn', { shape: s.id, at: ci }) },
                            { label: 'Insert column right', run: () => act('tableColumn', { shape: s.id, at: ci + 1 }) },
                            '-',
                            { label: 'Delete row', run: () => act('tableRow', { shape: s.id, at: ri, remove: true }) },
                            { label: 'Delete column', run: () => act('tableColumn', { shape: s.id, at: ci, remove: true }) },
                          ]);
                        }}
                        title="Double-click to edit this cell — right-click for rows and columns"
                      />
                    )))
                  )}
                  {/* Several selected: one dashed frame around all of them, no resize handles of its own. */}
                  {selectedIds.length > 1
                    ? (() => {
                        const boxes = selectedIds
                          .map((id) => (drag?.id === id ? drag.g : groupDrag?.boxes?.get(id)) ?? slide.shapes.find((s) => s.id === id)?.geometry)
                          .filter(Boolean);
                        if (!boxes.length) return null;
                        const x = Math.min(...boxes.map((b) => b.x));
                        const y = Math.min(...boxes.map((b) => b.y));
                        const r = Math.max(...boxes.map((b) => b.x + b.w));
                        const b2 = Math.max(...boxes.map((b) => b.y + b.h));
                        return <div className="sl-selection-frame" style={{ left: x, top: y, width: r - x, height: b2 - y, borderWidth: 1.5 / scale }} />;
                      })()
                    : null}
                  {selectedIds.length === 1 && selectedShape?.geometry && !selectedShape.hidden && !editing
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
                  {/* The rotation handle: floats above the selection, orbiting with it as it turns. */}
                  {selectedIds.length === 1 && selectedShape?.geometry && !selectedShape.hidden && !editing
                    ? (() => {
                        const g = drag?.id === selectedShape.id ? drag.g : selectedShape.geometry;
                        const cx = g.x + g.w / 2;
                        const cy = g.y + g.h / 2;
                        const rad = ((g.rot || 0) * Math.PI) / 180;
                        const gap = ROTATE_HANDLE_GAP + g.h / 2;
                        // The stem runs from the top edge to the handle, turning with the
                        // shape; drawn from the centre it cut through the shape itself.
                        const ex = cx + (g.h / 2) * Math.sin(rad);
                        const ey = cy - (g.h / 2) * Math.cos(rad);
                        const hx = cx + gap * Math.sin(rad);
                        const hy = cy - gap * Math.cos(rad);
                        const size = 10 / scale;
                        return (
                          <React.Fragment>
                            <div
                              className="sl-rotate-line"
                              style={{ left: ex, top: ey, width: ROTATE_HANDLE_GAP, height: 1.5 / scale, transformOrigin: '0 50%', transform: `rotate(${(g.rot || 0) - 90}deg)` }}
                            />
                            <div
                              className="sl-rotate-handle"
                              style={{ left: hx - size / 2, top: hy - size / 2, width: size, height: size }}
                              title="Drag to rotate the shape — hold Shift to snap to 15°"
                              onMouseDown={(e) => startRotate(e, selectedShape)}
                            />
                          </React.Fragment>
                        );
                      })()
                    : null}
                  {/*
                    The Animations tab's numbers beside each animated shape,
                    PowerPoint's own little tags: the click that starts the
                    effect (0 for one that starts with the slide), stacked
                    down the shape's left edge when it has several. A click
                    on one picks that effect.
                  */}
                  {tab === 'animations' && animations.length && !preview
                    ? (() => {
                        const byShape = new Map();
                        const size = 17 / scale;
                        return animations.map((e) => {
                          const g = slide.shapes.find((s) => String(s.id) === String(e.shapeId))?.geometry;
                          if (!g) return null;
                          const k = byShape.get(e.shapeId) || 0;
                          byShape.set(e.shapeId, k + 1);
                          return (
                            <button
                              key={`anim-${e.index}`}
                              type="button"
                              className={`sl-anim-badge sl-an-${e.kind}${currentAnim?.index === e.index ? ' current' : ''}`}
                              data-anim={e.index}
                              data-shape={e.shapeId}
                              style={{ left: g.x - size - 3 / scale, top: g.y + k * (size + 2 / scale), width: size, height: size, fontSize: 10.5 / scale, borderRadius: 3 / scale, borderWidth: 1 / scale }}
                              title={`${e.group} — ${e.name}${e.kind === 'exit' ? ' (exit)' : e.kind === 'emph' ? ' (emphasis)' : ''}, ${TRIGGER_WORDS[e.trigger] || e.trigger}`}
                              onMouseDown={(ev) => ev.stopPropagation()}
                              onClick={(ev) => { ev.stopPropagation(); act('animSelect', e.index); }}
                            >
                              {e.group}
                            </button>
                          );
                        });
                      })()
                    : null}
                  {/* Review → Comments: a marker where each thread on this slide is anchored. */}
                  {!masterPart && slideComments.length && !preview
                    ? markerSpots(slideComments, slide.shapes).map(({ thread, x, y }) => {
                        const size = 26 / scale;
                        return (
                          <button
                            key={`cm-${thread.id}`}
                            type="button"
                            className={`sl-cm-marker${thread.status === 'resolved' ? ' resolved' : ''}${commentSel === thread.id ? ' current' : ''}`}
                            data-comment={thread.id}
                            style={{ left: Math.max(0, Math.min(x, model.size.width - size)), top: Math.max(0, y), width: size, height: size, fontSize: 10 / scale, background: personColour(thread.author), borderWidth: 2 / scale }}
                            data-tip={`${thread.author}: ${thread.text.slice(0, 80)}${thread.replies.length ? ` (${thread.replies.length} repl${thread.replies.length === 1 ? 'y' : 'ies'})` : ''}`}
                            onMouseDown={(e) => e.stopPropagation()}
                            onClick={(e) => { e.stopPropagation(); act('selectComment', thread.id); }}
                          >
                            {thread.initials}
                            {thread.replies.length ? <span className="sl-cm-more">{thread.replies.length + 1}</span> : null}
                          </button>
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
                        if (editing.row != null) {
                          const box = s?.table?.cells?.[editing.row]?.[editing.col]?.box;
                          return box ? { left: box.x, top: box.y, width: box.w, height: box.h } : {};
                        }
                        return s?.geometry ? { left: s.geometry.x, top: s.geometry.y, width: s.geometry.w, height: s.geometry.h } : {};
                      })()}
                      onBlur={(e) => (editing.row != null ? commitTableCell(editing.id, editing.row, editing.col, e.target.value) : commitText(editing.id, e.target.value))}
                      onKeyDown={(e) => {
                        if (e.key === 'Escape') { setEditing(null); return; }
                        if (editing.row == null) {
                          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commitText(editing.id, e.target.value);
                          return;
                        }
                        // A table cell: Enter commits (Shift+Enter is a new
                        // line), Tab commits and moves on to the next cell —
                        // PowerPoint's own keys.
                        if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitTableCell(editing.id, editing.row, editing.col, e.target.value); return; }
                        if (e.key === 'Tab') {
                          e.preventDefault();
                          const s = slide.shapes.find((x) => x.id === editing.id);
                          const cols = s?.table?.cols || 1;
                          const rows = s?.table?.rows || 1;
                          let row = editing.row;
                          let col = editing.col + (e.shiftKey ? -1 : 1);
                          if (col >= cols) { col = 0; row += 1; }
                          if (col < 0) { col = cols - 1; row -= 1; }
                          const value = e.target.value;
                          commitTableCell(editing.id, editing.row, editing.col, value).then((next) => {
                            if (row < 0 || row >= rows) return;
                            const ns = next?.model?.slide?.shapes?.find((x) => x.id === editing.id);
                            const cell = ns?.table?.cells?.[row]?.[col];
                            setEditing({ id: editing.id, row, col, text: (cell?.paragraphs || []).map((p) => p.plain).join('\n') });
                          });
                        }
                      }}
                    />
                  ) : null}
                </div>
                </div>
              ) : (
                <Empty icon="slides" title="This presentation has no slides" />
              )}
            </div>
            {slide && !masterPart && view.mode === 'notes' ? (
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
              title={view.pane === 'layers' ? 'Layers' : view.pane === 'designs' ? 'Designs' : view.pane === 'animations' ? 'Animation Pane' : view.pane === 'comments' ? 'Comments' : 'Format'}
              actions={<Button icon="close" title="Close the pane" onClick={() => act('pane', view.pane)} />}
            >
              {view.pane === 'comments' ? (
                <CommentsPane
                  threads={slideComments}
                  draft={commentDraft}
                  selected={commentSel}
                  shapes={slide?.shapes || []}
                  me={model.me || 'You'}
                  onSelect={(id) => setCommentSel(id)}
                  onNew={() => act('newComment')}
                  onPost={(text) => act('postComment', text)}
                  onCancelDraft={() => setCommentDraft(null)}
                  onReply={(thread, text) => act('replyComment', { thread, text })}
                  onResolve={(thread, resolved) => act('resolveComment', { thread, resolved })}
                  onDelete={(thread) => act('deleteComment', thread)}
                  onDeleteReply={(thread, reply) => act('deleteReply', { thread, reply })}
                />
              ) : view.pane === 'layers' ? (
                <LayersPane slide={slide} selected={selected} selectedIds={selectedIds} onSelect={setSelected} onToggle={toggleSelected} act={act} />
              ) : view.pane === 'animations' ? (
                <AnimationPane slide={slide} current={currentAnim} act={act} playing={preview?.kind === 'animation'} />
              ) : view.pane === 'designs' ? (
                <DesignsPane layouts={model.layouts} current={slide?.layout || null} size={model.size} act={act} />
              ) : (
                <FormatPane shape={selectedShape} theme={slide?.theme} act={act} />
              )}
            </Panel>
          ) : null}
          {review.pane && model ? (
            <Panel right width={300} resizable title={review.paneTitle} actions={<Button icon="close" title="Close the pane" onClick={review.close} />}>
              {review.paneNode}
            </Panel>
          ) : null}
          {menu.node}
        </>
      )}

      {review.dialogs}

      {shortcutsOpen ? <SlidesShortcutsDialog onClose={() => setShortcutsOpen(false)} /> : null}

      {equationOpen ? (
        <EquationDialog
          initial={equationOpen.initial || ''}
          display={equationOpen.display !== false}
          editing={equationOpen.shape != null}
          onClose={() => setEquationOpen(null)}
          onInsert={(linear, display) => act('putEquation', { linear, display })}
        />
      ) : null}

      {partRename ? (
        <PartNameDialog
          name={partRename.name}
          kind={partRename.kind}
          onClose={() => setPartRename(null)}
          onApply={async (name) => {
            const next = await apply({ op: 'renamePart', part: partRename.part, name });
            if (next) setPartRename(null);
          }}
        />
      ) : null}

      {designOpen && doc ? (
        <DesignGallery
          key={designOpen.kind}
          kind={designOpen.kind}
          anchor={designOpen.anchor}
          shell={shell}
          docId={doc.id}
          slide={index}
          onClose={() => setDesignOpen(null)}
          onPick={(it) => {
            if (designOpen.kind === 'themes') act('applyTheme', it);
            else if (designOpen.kind === 'variants') act('applyVariant', it.id);
            else if (designOpen.kind === 'colours') act('themeColors', it.id);
            else if (designOpen.kind === 'fonts') act('themeFonts', it.id);
            else if (designOpen.kind === 'effects') act('themeEffects', it.id);
          }}
          onCustomise={(info) => act(designOpen.kind === 'colours' ? 'customColours' : 'customFonts', info)}
        />
      ) : null}

      {customColours ? (
        <CustomColoursDialog
          info={customColours}
          onClose={() => setCustomColours(null)}
          onSave={async (colors, name) => {
            const next = await apply({ op: 'setThemeColors', colors, name });
            if (next) { setCustomColours(null); toast(`Theme colours "${name}" saved in this deck`, { tone: 'good' }); }
          }}
        />
      ) : null}

      {customFonts ? (
        <CustomFontsDialog
          info={customFonts}
          onClose={() => setCustomFonts(null)}
          onSave={async (fonts, name) => {
            const next = await apply({ op: 'setThemeFonts', major: fonts.major, minor: fonts.minor, name });
            if (next) { setCustomFonts(null); toast(`Theme fonts "${name}" saved in this deck`, { tone: 'good' }); }
          }}
        />
      ) : null}

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

      {chartDataOpen != null ? (() => {
        const chartShape = slide?.shapes?.find((s) => s.id === chartDataOpen);
        if (!chartShape?.chart) return null;
        return (
          <ChartDataDialog
            chart={chartShape.chart}
            onClose={() => setChartDataOpen(null)}
            onApply={async (data) => {
              await apply({ op: 'setChartData', slide: index, shape: chartDataOpen, ...data });
              setChartDataOpen(null);
            }}
          />
        );
      })() : null}

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

      {gate.node}
      {protection.node}
    </AppFrame>
  );
}

/** How far above the shape's top edge the rotation handle floats, in slide pixels. */
const ROTATE_HANDLE_GAP = 26;

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
/** Shape Effects: PowerPoint's offset shadows, in points and degrees, and none — each the `shadow` key alone, merged in beside whatever other effects the shape already has. */
const SHADOW_PRESETS = {
  none: null,
  br: { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 },
  b: { dist: 3, dir: 90, blur: 4, color: '#000000', alpha: 0.4 },
  r: { dist: 3, dir: 0, blur: 4, color: '#000000', alpha: 0.4 },
  tl: { dist: 3, dir: 225, blur: 4, color: '#000000', alpha: 0.4 },
  c: { dist: 0, dir: 0, blur: 6, color: '#000000', alpha: 0.45 },
};
/** Shape Effects → Glow: PowerPoint's own gallery radii, in points. */
const GLOW_SIZES = [5, 8, 11, 18];
/** Shape Effects → Reflection: the three gallery presets, by name. */
const REFLECTION_LABELS = [['tight', 'Tight Reflection'], ['half', 'Half Reflection'], ['full', 'Full Reflection'], [null, 'No Reflection']];
const SHADOW_LABELS = [['none', 'None'], ['br', 'Bottom right'], ['b', 'Below'], ['r', 'Right'], ['tl', 'Top left'], ['c', 'All round']];
/**
 * The scene reads a shape's effects back in its own units — a shadow's
 * distance and blur in pixels, a glow's and a soft edge's radius in points
 * but named `radiusPt` — while `setShapeStyle` always writes them in
 * points, under the plain names (`dist`, `blur`, `radius`) the Format pane
 * and Shape Effects menu use for a *new* value. Merging a shape's existing
 * effects into one more change (`shapeEffects`/`shapeShadow`, below) has to
 * go through here first, or a preserved effect quietly loses whatever the
 * two sides name differently — which is exactly how asking for a glow kept
 * a soft edge the shape already had, in name only: written back with no
 * radius the writer recognised, so it silently fell to the writer's own
 * default instead of the value on screen.
 */
const PX_PER_PT = 96 / 72;
function effectsToWriteSpec(effects) {
  if (!effects) return {};
  const sh = effects.shadow;
  const glow = effects.glow;
  const softEdge = effects.softEdge;
  return {
    shadow: sh ? { dist: sh.distPx / PX_PER_PT, dir: sh.dir, blur: sh.blurPx / PX_PER_PT, color: sh.color, alpha: sh.alpha } : null,
    glow: glow ? { radius: glow.radiusPt, color: glow.color, alpha: glow.alpha } : undefined,
    softEdge: softEdge ? { radius: softEdge.radiusPt } : undefined,
    reflection: effects.reflection || undefined,
  };
}

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
  if (shape.kind === 'group') return <div className="sl-pane-empty">A group has no fill or outline of its own — format a shape inside it.</div>;
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
  // The fill's own kind — what the segmented control shows current, and
  // which of the four sub-panels below it draws.
  const fillType = shape.fill?.type === 'gradient' ? 'gradient' : shape.fill?.type === 'picture' ? 'picture' : shape.fill?.type === 'none' ? 'none' : 'solid';
  const transparency = shape.fill?.type === 'solid' ? Math.round((1 - (shape.fill.alpha ?? 1)) * 100) : 0;
  const angle = shape.fill?.type === 'gradient' ? Math.round(shape.fill.angle ?? 90) : 90;
  const glow = shape.effects?.glow || null;
  const softEdge = shape.effects?.softEdge || null;
  const reflection = shape.effects?.reflection || null;

  return (
    <div className="sl-format">
      <div className="sl-format-head">{shape.name || shape.kind}</div>
      <section className="sl-format-fill">
        <h4>Fill</h4>
        <div className="sl-format-row sl-format-kinds">
          {[['none', 'No fill'], ['solid', 'Solid'], ['gradient', 'Gradient'], ['picture', 'Picture']].map(([key, label]) => (
            <button
              key={key}
              type="button"
              className={`sl-chip${fillType === key ? ' current' : ''}`}
              onClick={() => {
                if (key === 'none') act('shapeFill', 'none');
                else if (key === 'solid') act('shapeFill', { color: fill && fill !== 'none' ? fill : (colours.accent1 || '#4472C4'), alpha: 1 });
                else if (key === 'gradient') act('shapeFill', { gradient: { preset: 'light', color: { scheme: 'accent1' } } });
                else act('pickPictureFill');
              }}
            >
              {label}
            </button>
          ))}
        </div>
        {fillType === 'solid' ? (
          <>
            {swatches((hex) => act('shapeFill', { color: hex, alpha: shape.fill?.alpha ?? 1 }), fill)}
            <div className="sl-format-row">
              <input type="color" className="sl-colour" title="Any colour" value={fill && fill !== 'none' ? fill : '#4472c4'} onChange={(e) => act('shapeFill', { color: e.target.value, alpha: shape.fill?.alpha ?? 1 })} />
              <label className="sl-format-slider">
                Transparency {transparency}%
                <input type="range" min="0" max="100" value={transparency} onChange={(e) => act('shapeFill', { color: fill && fill !== 'none' ? fill : '#4472c4', alpha: 1 - Number(e.target.value) / 100 })} />
              </label>
            </div>
          </>
        ) : null}
        {fillType === 'gradient' ? (
          <div className="sl-format-row" style={{ flexWrap: 'wrap' }}>
            <button type="button" className="sl-chip" onClick={() => act('shapeFill', { gradient: { preset: 'light', color: { scheme: 'accent1' }, angle } })}>Light Variation</button>
            <button type="button" className="sl-chip" onClick={() => act('shapeFill', { gradient: { preset: 'dark', color: { scheme: 'accent1' }, angle } })}>Dark Variation</button>
            <label className="sl-format-slider">
              Angle {angle}°
              <input type="range" min="0" max="360" value={angle} onChange={(e) => act('shapeFill', { gradient: { stops: shape.fill?.stops?.map((s) => ({ pos: s.offset, color: s.color, alpha: s.alpha })), angle: Number(e.target.value) } })} />
            </label>
          </div>
        ) : null}
        {fillType === 'picture' && shape.fill?.embed ? (
          <div className="sl-format-row">
            <Button icon="picture" label="Choose picture…" onClick={() => act('pickPictureFill')} />
            <label className="sl-format-check">
              <input type="checkbox" checked={Boolean(shape.fill.tile)} onChange={(e) => act('shapeFill', { picture: { tile: e.target.checked } })} /> Tile
            </label>
          </div>
        ) : null}
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
      <section className="sl-format-glow">
        <h4>Glow</h4>
        <div className="sl-format-row" style={{ flexWrap: 'wrap' }}>
          <button type="button" className={`sl-chip${!glow ? ' current' : ''}`} onClick={() => act('shapeEffects', { glow: null })}>No glow</button>
          {GLOW_SIZES.map((pt) => (
            <button key={pt} type="button" className={`sl-chip${glow && Math.round(glow.radiusPt) === pt ? ' current' : ''}`} onClick={() => act('shapeEffects', { glow: { radius: pt, color: glow?.color || colours.accent1 || '#4472C4', alpha: glow?.alpha ?? 0.6 } })}>{pt} pt</button>
          ))}
          <input type="color" className="sl-colour" title="Glow colour" value={glow?.color || colours.accent1 || '#4472c4'} onChange={(e) => act('shapeEffects', { glow: { radius: glow?.radiusPt || 8, color: e.target.value, alpha: glow?.alpha ?? 0.6 } })} />
        </div>
      </section>
      <section className="sl-format-softedge">
        <h4>Soft Edges</h4>
        <div className="sl-format-row" style={{ flexWrap: 'wrap' }}>
          <button type="button" className={`sl-chip${!softEdge ? ' current' : ''}`} onClick={() => act('shapeEffects', { softEdge: null })}>None</button>
          {[1, 2.5, 5, 10].map((pt) => (
            <button key={pt} type="button" className={`sl-chip${softEdge && Math.round(softEdge.radiusPt * 2) === Math.round(pt * 2) ? ' current' : ''}`} onClick={() => act('shapeEffects', { softEdge: { radius: pt } })}>{pt} pt</button>
          ))}
        </div>
      </section>
      <section className="sl-format-reflection">
        <h4>Reflection</h4>
        <div className="sl-format-row" style={{ flexWrap: 'wrap' }}>
          {REFLECTION_LABELS.map(([key, label]) => (
            <button key={label} type="button" className={`sl-chip${reflection === key ? ' current' : ''}`} onClick={() => act('shapeEffects', { reflection: key })}>{label}</button>
          ))}
        </div>
      </section>
    </div>
  );
}

/** The icon a layer row shows for its shape. */
/** How the pane and the numbers on the stage say each start. */
const TRIGGER_WORDS = { onClick: 'on click', withPrevious: 'with the previous', afterPrevious: 'after the previous' };
const KIND_WORDS = { entr: 'Entrance', emph: 'Emphasis', exit: 'Exit', path: 'Motion path', media: 'Media', other: 'Effect' };

/**
 * Animations → Animation Pane: the slide's sequence as PowerPoint lists it —
 * each effect numbered by the click that starts it, a mouse for On Click, a
 * clock for After Previous, a star in the kind's colour, the shape and the
 * effect. Click a row to pick it; drag it (or use the arrows) to move it;
 * right-click for its start; the bin takes it out; Play All plays the lot
 * on the stage.
 */
function AnimationPane({ slide, current, act, playing }) {
  const list = slide?.animations || [];
  const names = new Map((slide?.shapes || []).map((s) => [String(s.id), s.name || `Shape ${s.id}`]));
  const [dragFrom, setDragFrom] = React.useState(null);
  const [over, setOver] = React.useState(null);
  const at = current?.index ?? -1;
  return (
    <div className="sl-animpane">
      <div className="sl-layers-tools">
        <Button icon="play" label={playing ? 'Playing…' : 'Play All'} className="sl-animpane-play" disabled={!list.length || playing} title={list.length ? "Play All — the slide's animations, in order, on the stage" :'Play All — this slide has no animations yet'} onClick={() => act('preview', 'animation')} />
        <Spacer />
        <Button icon="chevronUp" className="sl-animpane-up" title="Move Earlier" disabled={at <= 0} onClick={() => act('animMove', 'earlier')} />
        <Button icon="chevronDown" className="sl-animpane-down" title="Move Later" disabled={at < 0 || at >= list.length - 1} onClick={() => act('animMove', 'later')} />
        <Button icon="trash" className="sl-animpane-remove" title="Remove — take the picked effect out" disabled={at < 0} onClick={() => act('animRemove', at)} />
      </div>
      {list.length ? (
        list.map((e, i) => {
          const newGroup = i === 0 || e.group !== list[i - 1].group;
          return (
            <div
              key={i}
              className={`sl-animrow sl-an-${e.kind}${i === at ? ' active' : ''}${over === i && dragFrom != null && dragFrom !== i ? ' drop' : ''}`}
              data-anim={i}
              draggable
              onDragStart={(ev) => { setDragFrom(i); ev.dataTransfer.effectAllowed = 'move'; ev.dataTransfer.setData('text/plain', String(i)); }}
              onDragOver={(ev) => { ev.preventDefault(); setOver(i); }}
              onDragLeave={() => setOver((o) => (o === i ? null : o))}
              onDrop={(ev) => { ev.preventDefault(); const from = dragFrom; setDragFrom(null); setOver(null); if (from != null && from !== i) act('animMove', { from, to: i }); }}
              onDragEnd={() => { setDragFrom(null); setOver(null); }}
              onClick={() => act('animSelect', i)}
              onContextMenu={(ev) => { act('animSelect', i); act('animMenu', { ev, index: i }); }}
              title={`${KIND_WORDS[e.kind] || 'Effect'}: ${e.name}${e.direction ? ` (${e.direction})` : ''} — starts ${TRIGGER_WORDS[e.trigger] || e.trigger}${e.delay ? `, ${e.delay} s later` : ''}; ${e.duration ? `${e.duration} s` : 'at once'}${e.known ? '' : '. Kept as the file has it.'}`}
            >
              <span className="sl-animrow-n">{newGroup ? e.group : ''}</span>
              <span className="sl-animrow-trigger">{e.trigger === 'onClick' ? <Icon name="mouse" size={13} /> : e.trigger === 'afterPrevious' ? <Icon name="clock" size={13} /> : null}</span>
              <Icon name="star" size={14} className="sl-animrow-star" />
              <span className="sl-layer-text">
                <span className="sl-layer-title">{names.get(String(e.shapeId)) || `Shape ${e.shapeId}`}</span>
                <span className="sl-layer-words">{e.name}{e.kind === 'exit' ? ' (exit)' : ''}{e.duration ? ` · ${e.duration.toFixed(2)} s` : ''}{e.delay ? ` · after ${e.delay.toFixed(2)} s` : ''}</span>
              </span>
            </div>
          );
        })
      ) : (
        <div className="sl-pane-empty">No animations on this slide. Select a shape and pick an effect from the Animations tab.</div>
      )}
    </div>
  );
}

const LAYER_ICONS = { picture: 'picture', table: 'table', chart: 'chart', connector: 'minus', shape: 'shape', unsupported: 'shape' };

/**
 * The slide's shapes as layers, top-most first — PowerPoint's selection pane
 * with the parts people use: a click selects the shape on the stage (Shift
 * or Ctrl adds or removes one), the eye hides it (it stays in the file,
 * undrawn), the arrows change the drawing order, a double-click renames it.
 * A group is one row, its own members listed indented directly under it —
 * the order the arrows change is the top level's, a group counting as one
 * shape there exactly as it does on the stage.
 */
function LayersPane({ slide, selected, selectedIds = [], onSelect, onToggle, act }) {
  const [renaming, setRenaming] = React.useState(null);
  const all = slide?.shapes || [];
  const topLevel = all.filter((s) => s.groupId == null);
  const rows = [...topLevel].reverse();
  const childrenOf = (id) => [...all.filter((s) => s.groupId === id)].reverse();
  const pos = selected != null ? topLevel.findIndex((s) => s.id === selected) : -1;
  const n = topLevel.length;
  const selectedShape = selected != null ? all.find((s) => s.id === selected) : null;

  const row = (s, indent) => (
    <div
      key={s.id}
      className={`sl-layer${selectedIds.includes(s.id) ? ' active' : ''}${s.hidden ? ' off' : ''}`}
      data-shape={s.id}
      style={indent ? { paddingLeft: 14 + indent * 16 } : undefined}
      onClick={(e) => ((e.shiftKey || e.ctrlKey || e.metaKey) && s.groupId == null ? onToggle(s.id) : onSelect(s.id))}
      onDoubleClick={() => setRenaming({ id: s.id, name: s.name || '' })}
      title="Click to select (Shift/Ctrl adds or removes one); double-click to rename"
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
      <Icon name={s.kind === 'group' ? 'grid' : s.text ? 'textbox' : LAYER_ICONS[s.kind] || 'shape'} size={14} />
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
  );

  return (
    <div className="sl-layers">
      <div className="sl-layers-tools">
        <Button icon="chevronUp" title="Bring forward" disabled={pos < 0 || pos >= n - 1} onClick={() => act('order', 'forward')} />
        <Button icon="chevronDown" title="Send backward" disabled={pos <= 0} onClick={() => act('order', 'backward')} />
        <Button label="Front" title="Bring to front" disabled={pos < 0 || pos >= n - 1} onClick={() => act('order', 'front')} />
        <Button label="Back" title="Send to back" disabled={pos <= 0} onClick={() => act('order', 'back')} />
        <Spacer />
        <Button icon="grid" title={selectedIds.length < 2 ? 'Select two or more shapes to group' : 'Group'} disabled={selectedIds.length < 2} onClick={() => act('group')} />
        <Button icon="grid" title={selectedShape?.kind !== 'group' ? 'Select a group to ungroup' : 'Ungroup'} disabled={selectedShape?.kind !== 'group'} onClick={() => act('ungroup')} />
        <Spacer />
        <Button icon="trash" title="Delete the selected shape(s)" disabled={!selectedIds.length} onClick={() => act('deleteShape')} />
      </div>
      {rows.length ? (
        rows.map((s) => (
          <React.Fragment key={s.id}>
            {row(s, 0)}
            {s.kind === 'group' ? childrenOf(s.id).map((c) => row(c, 1)) : null}
          </React.Fragment>
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
.sl-format-kinds { flex-wrap: wrap; gap: 6px; }
.sl-chip { white-space: nowrap; border: 1px solid var(--line); background: var(--surface); border-radius: 999px; padding: 3px 10px; font-size: 11.5px; color: var(--ink-2); }
.sl-chip:hover, .sl-chip.current { color: var(--accent); border-color: var(--accent); }
.sl-format-slider { min-width: 120px; }
.sl-format-slider input[type="range"] { width: 100%; }
.sl-format-check { flex-direction: row !important; align-items: center; gap: 5px !important; }
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
/* Slide Master view's strip: the master, then its layouts indented under it, smaller, joined by a rule. */
.sl-masterstrip { gap: 8px; }
.sl-mthumb { flex-direction: column; gap: 3px; position: relative; }
.sl-mthumb .sl-thumb-card { flex: none; width: 100%; }
.sl-thumb.sl-mthumb-layout { padding-left: 22px; }
.sl-mthumb-layout::before { content: ''; position: absolute; left: 9px; top: -8px; bottom: 50%; width: 9px; border-left: 1px solid var(--line-strong); border-bottom: 1px solid var(--line-strong); border-bottom-left-radius: 4px; }
.sl-mthumb-name { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--ink-2); padding: 0 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-mthumb-master .sl-mthumb-name { font-weight: 600; color: var(--ink); }
.sl-mthumb.active .sl-mthumb-name { color: var(--accent); }
.sl-mthumb-used { flex: none; min-width: 16px; height: 16px; padding: 0 4px; border-radius: 8px; background: var(--surface-2); border: 1px solid var(--line-soft); color: var(--ink-3); font-size: 10px; display: inline-grid; place-items: center; }
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
/* Slide Show → Hide Slide: left dim, the way PowerPoint fades a hidden slide, its number struck through. */
.sl-thumb.hidden .sl-thumb-card, .sl-sortercard.hidden .sl-thumb-pic, .sl-sortercard.hidden .sl-thumb-title { opacity: .45; }
.sl-thumb.hidden .sl-thumb-n, .sl-sortercard.hidden .sl-sortern { text-decoration: line-through; }
.sl-fit { position: relative; flex: none; }

.sl-stage { position: relative; flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center; padding: 22px; background: var(--window); }
.sl-slide { position: relative; box-shadow: var(--shadow-2); background: #fff; }
.sl-svg svg { display: block; width: 100%; height: 100%; }
.sl-hit { position: absolute; border: 1px solid transparent; background: transparent; border-radius: 2px; min-height: 8px; min-width: 8px; cursor: move; }
.sl-hit.dragging { border: 1px dashed var(--accent); background: rgba(43, 95, 217, 0.08); z-index: 4; }
.sl-handle { position: absolute; z-index: 5; background: #fff; border: 1.5px solid var(--accent); border-radius: 2px; box-sizing: border-box; }
.sl-stage:focus { outline: none; }

.sl-hit:hover { border-color: var(--accent-line); background: rgba(43, 95, 217, 0.06); }
/* Home → Find (Ctrl+F) and Replace (Ctrl+H): a small pane pinned to the
   stage's own top right, over whatever slide is on screen — Word's find
   pane, not a dialog that covers the slide. It sits above the stage's
   scroll, the way an absolutely positioned child of a relatively positioned
   scroller always does. */
.sl-find {
  position: absolute; top: 12px; right: 22px; z-index: 20; display: flex; flex-direction: column; gap: 6px;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3);
  padding: 8px 10px; box-shadow: var(--shadow-2); width: 268px;
}
.sl-find-row { display: flex; align-items: center; gap: 6px; }
.sl-find-row .rw-input { flex: 1; min-width: 0; }
.sl-find-count { font-size: 12px; color: var(--ink-3); white-space: nowrap; min-width: 3.5em; text-align: right; }
.sl-find-case { display: flex; align-items: center; gap: 5px; font-size: 12px; color: var(--ink-2); white-space: nowrap; }
.sl-find-note { font-size: 12px; color: var(--ink-3); }
/* The shape the pane is on now, over its ordinary selection ring, so
   stepping through hits reads as something distinct from a plain click. */
.sl-hit.find-current { border-color: #d9a300; box-shadow: 0 0 0 2px rgba(217, 163, 0, 0.35); }
/* A table's per-cell hit layer, over the grid the SVG already drew. */
.sl-cell-hit { position: absolute; z-index: 2; cursor: text; border: 1px solid transparent; box-sizing: border-box; }
.sl-cell-hit:hover { border-color: var(--accent-line); background: rgba(43, 95, 217, 0.06); }
.sl-editor {
  position: absolute; border: 2px solid var(--accent); border-radius: 3px; padding: 4px 6px;
  font: inherit; font-size: 15px; background: #fff; color: #111; resize: none; outline: none; z-index: 5;
}
.sl-notes {
  border-top: 1px solid var(--line); background: var(--chrome); padding: 9px 18px;
  font-size: 12.5px; color: var(--ink-2); max-height: 110px; overflow: auto; white-space: pre-wrap;
}

.sl-hit.selected { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
/* The primary of a multi-selection — the one Font/Paragraph act on — wears
   a solid ring; the rest of the selection wears the ordinary dashed one. */
.sl-hit.selected.primary { box-shadow: 0 0 0 2px var(--accent); }
/* Several selected: one dashed frame around the whole selection, under the
   individual outlines and never itself a drag or resize target. */
.sl-selection-frame { position: absolute; z-index: 3; pointer-events: none; border: 1.5px dashed var(--accent); box-sizing: border-box; }
/* The rotation handle: a small circle above the selection, joined to it by
   a thin line that orbits with the shape as it turns. */
.sl-rotate-line { position: absolute; z-index: 5; background: var(--accent); pointer-events: none; }
.sl-rotate-handle {
  position: absolute; z-index: 6; background: #fff; border: 1.5px solid var(--accent); border-radius: 50%;
  box-sizing: border-box; cursor: grab;
}
.sl-rotate-handle:active { cursor: grabbing; }
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
/* The show: two layers in one box while a transition plays, one otherwise. */
.sl-show-stage { position: relative; overflow: hidden; background: #000; isolation: isolate; }
.sl-show-layer { position: absolute; inset: 0; background: #fff; will-change: transform, opacity, clip-path; }
.sl-show-layer svg { display: block; width: 100%; height: 100%; }
/* Animations: a shape turns and grows about its own middle. */
.sl-show-layer g[data-shape], .sl-preview g[data-shape], .pv-stage g[data-shape], .pv-thumb g[data-shape] { transform-box: fill-box; transform-origin: 50% 50%; }
/* The numbers beside animated shapes on the stage (Animations tab): PowerPoint's grey tags, the picked one in the accent. */
.sl-anim-badge { position: absolute; z-index: 7; display: grid; place-items: center; padding: 0; border: 1px solid #8a8f98; background: #f3f4f6; color: #30343b; font-weight: 600; font-variant-numeric: tabular-nums; line-height: 1; cursor: pointer; box-sizing: border-box; box-shadow: 0 1px 2px rgba(0,0,0,.12); }
.sl-anim-badge:hover { border-color: var(--accent); color: var(--accent); }
.sl-anim-badge.current { background: var(--accent); border-color: var(--accent); color: #fff; }
/* Animations → Animation Pane. */
.sl-animpane { display: flex; flex-direction: column; }
.sl-animrow { display: flex; align-items: center; gap: 7px; padding: 6px 10px 6px 6px; border-bottom: 1px solid var(--line-soft); font-size: 12.5px; color: var(--ink); cursor: default; }
.sl-animrow:hover { background: var(--surface-2); }
.sl-animrow.active { background: var(--selected); box-shadow: inset 3px 0 0 var(--accent); }
.sl-animrow.drop { box-shadow: inset 0 2px 0 var(--accent); }
.sl-animrow-n { width: 16px; flex: none; text-align: right; font-size: 11.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.sl-animrow-trigger { width: 14px; flex: none; display: grid; place-items: center; color: var(--ink-3); }
.sl-animrow-star { flex: none; }
/* A star in each kind's own colour, as PowerPoint draws them: green in, gold emphasis, red out. */
.sl-an-entr .sl-animrow-star, .rw-btn.sl-an-entr svg { color: #2e9a4a; }
.sl-an-emph .sl-animrow-star, .rw-btn.sl-an-emph svg { color: #c28d00; }
.sl-an-exit .sl-animrow-star, .rw-btn.sl-an-exit svg { color: #cf4338; }
.sl-an-path .sl-animrow-star, .sl-an-media .sl-animrow-star, .sl-an-other .sl-animrow-star, .rw-btn.sl-an-none svg { color: var(--ink-3); }
.sl-animrow-star path, .rw-btn.sl-an-entr svg path, .rw-btn.sl-an-emph svg path, .rw-btn.sl-an-exit svg path { fill: color-mix(in srgb, currentColor 22%, transparent); }
/* Transitions → Preview, over the editing stage. */
.sl-preview { position: absolute; inset: 0; z-index: 9; overflow: hidden; background: #000; pointer-events: none; }
.sl-preview-layer { position: absolute; inset: 0; background: #fff; }
.sl-preview-layer:empty { background: #000; }
.sl-preview-layer svg { display: block; width: 100%; height: 100%; }
/* The strip's star: this slide has a transition (or animations). */
.sl-thumb-n { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; }
.sl-thumb-fx { display: grid; place-items: center; color: var(--ink-3); text-decoration: none; }
.sl-thumb.active .sl-thumb-fx { color: var(--accent); }
.sl-sorterfx { position: absolute; right: 6px; bottom: 4px; display: grid; place-items: center; color: var(--ink-3); background: rgba(255,255,255,.85); padding: 1px 3px; border-radius: 3px; }
/* Ribbon fields: Duration, the Advance Slide ticks and After. */
.sl-rb-field { display: flex; align-items: center; gap: 6px; height: 26px; font-size: 12px; color: var(--ink-2); white-space: nowrap; padding: 0 4px; }
.sl-rb-field svg { color: var(--ink-2); }
.sl-rb-check { display: flex; align-items: center; gap: 6px; }
.sl-rb-field input[type="checkbox"] { margin: 0; accent-color: var(--accent); }
.sl-rb-seconds { width: 64px; height: 24px; padding: 0 4px 0 7px; font-size: 12px; font-variant-numeric: tabular-nums; }
.sl-rb-label { min-width: 54px; }
.sl-rb-caption { height: 20px; display: flex; align-items: center; padding: 0 4px; font-size: 11px; color: var(--ink-3); }
.sl-present-bar {
  position: fixed; z-index: 1; bottom: 14px; left: 50%; transform: translateX(-50%);
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
/** Slide Master → Rename: a layout's or the master's name. */
function PartNameDialog({ name: current, kind, onClose, onApply }) {
  const [name, setName] = useState(current || '');
  const label = kind === 'master' ? 'Rename master' : 'Rename layout';
  return (
    <Dialog
      title={label}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Rename" className="sl-partname-ok" disabled={!name.trim()} onClick={() => onApply(name.trim())} />
        </>
      }
    >
      <Field label={kind === 'master' ? 'Master name' : 'Layout name'}>
        <input className="rw-input sl-partname" autoFocus value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && name.trim()) onApply(name.trim()); }} />
      </Field>
    </Dialog>
  );
}

/** What only a slide can take: refused while Slide Master view is on the stage. */
const MASTER_REFUSED = new Set(['addComment', 'replyComment', 'resolveComment', 'removeComment', 'removeAllComments', 'insertSlide', 'duplicateSlide', 'removeSlide', 'moveSlide', 'setNotes', 'setTransition', 'applyTransitionToAll', 'addAnimation', 'setAnimation', 'removeAnimation', 'moveAnimation', 'removeShapeAnimations', 'setSlideHidden', 'addSection', 'renameSection', 'removeSection', 'removeAllSections', 'applyLayout', 'resetSlide', 'setFooter', 'replaceText', 'replaceHit', 'replaceAllHits', 'addTable', 'addChart', 'setChartData', 'setTableCell', 'insertTableRow', 'removeTableRow', 'insertTableColumn', 'removeTableColumn']);

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

/** Insert → Chart → double-click, or Edit Data…: the categories, the series and their values, rewritten into the chart part. */
function ChartDataDialog({ chart, onClose, onApply }) {
  const [type, setType] = useState(chart.type || 'column');
  const [title, setTitle] = useState(chart.title || '');
  const [categories, setCategories] = useState([...(chart.categories || [])]);
  const [series, setSeries] = useState((chart.series || []).map((s) => ({ name: s.name || '', values: [...(s.values || [])] })));

  const setCategory = (row, value) => setCategories((c) => c.map((v, i) => (i === row ? value : v)));
  const setSeriesName = (col, value) => setSeries((list) => list.map((s, i) => (i === col ? { ...s, name: value } : s)));
  const setValue = (row, col, value) => setSeries((list) => list.map((s, i) => (i === col ? { ...s, values: s.values.map((v, ri) => (ri === row ? value : v)) } : s)));

  const addRow = () => {
    setCategories((c) => [...c, `Category ${c.length + 1}`]);
    setSeries((list) => list.map((s) => ({ ...s, values: [...s.values, 0] })));
  };
  const removeRow = (row) => {
    setCategories((c) => c.filter((_, i) => i !== row));
    setSeries((list) => list.map((s) => ({ ...s, values: s.values.filter((_, i) => i !== row) })));
  };
  const addColumn = () => setSeries((list) => [...list, { name: `Series ${list.length + 1}`, values: categories.map(() => 0) }]);
  const removeColumn = (col) => setSeries((list) => list.filter((_, i) => i !== col));

  const cell = { padding: '2px 4px' };
  return (
    <Dialog
      title="Edit chart data"
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button
            primary
            label="Apply"
            className="sl-chart-apply"
            disabled={!categories.length || !series.length}
            onClick={() => onApply({
              type,
              title: title.trim() || null,
              categories,
              series: series.map((s) => ({ name: s.name, values: s.values.map((v) => (v === '' || v == null ? null : Number(v))) })),
            })}
          />
        </>
      }
    >
      <div className="ml-form sl-chart-data">
        <div style={{ display: 'flex', gap: 12 }}>
          <div style={{ flex: '0 0 160px' }}>
            <Field label="Chart type">
              <Select className="sl-chart-type" value={type} onChange={(e) => setType(e.target.value)}>
                <option value="column">Column</option>
                <option value="bar">Bar</option>
                <option value="line">Line</option>
                <option value="pie">Pie</option>
              </Select>
            </Field>
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Title">
              <input className="rw-input sl-chart-title" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="No title" />
            </Field>
          </div>
        </div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <thead>
              <tr>
                <th style={cell} />
                {series.map((s, ci) => (
                  <th key={ci} style={cell}>
                    <input className="rw-input sl-chart-series-name" data-col={ci} value={s.name} onChange={(e) => setSeriesName(ci, e.target.value)} style={{ width: 90 }} />
                  </th>
                ))}
                <th style={{ ...cell, width: 28 }}>
                  <Button icon="plus" title="Add a series" onClick={addColumn} />
                </th>
              </tr>
            </thead>
            <tbody>
              {categories.map((catName, ri) => (
                <tr key={ri}>
                  <td style={cell}>
                    <input className="rw-input sl-chart-category" data-row={ri} value={catName} onChange={(e) => setCategory(ri, e.target.value)} style={{ width: 90 }} />
                  </td>
                  {series.map((s, ci) => (
                    <td key={ci} style={cell}>
                      <input
                        className="rw-input sl-chart-cell"
                        data-row={ri}
                        data-col={ci}
                        value={s.values[ri] ?? ''}
                        onChange={(e) => setValue(ri, ci, e.target.value)}
                        style={{ width: 70 }}
                      />
                    </td>
                  ))}
                  <td style={cell}>
                    <Button icon="trash" title="Remove this category" disabled={categories.length <= 1} onClick={() => removeRow(ri)} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div>
          <Button icon="plus" label="Category" onClick={addRow} />
          {series.length > 1 ? <Button icon="trash" label="Remove last series" onClick={() => removeColumn(series.length - 1)} /> : null}
        </div>
        <p className="rw-hint" style={{ margin: 0 }}>A pie chart draws only its first series.</p>
      </div>
    </Dialog>
  );
}

/** Home → Editing: the words found on every slide, each match a step to its shape, and replaced across the deck. */
/**
 * Home → Find (Ctrl+F) and Replace (Ctrl+H): a small pane at the stage's own
 * top right, the way Word's find pane sits over the page rather than
 * covering it. Hits are recomputed as the words typed change (a short pause
 * so a fast typist is not chased by a search on every letter); Next and
 * Previous walk them one at a time, each landing on the hit's slide and
 * shape. `mode` decides whether the replace row shows from the start —
 * Ctrl+F opens to find only, Ctrl+H with it open — but either can reach the
 * other's row once open.
 */
function FindPane({ mode, onClose, onSearch, onGoto, onReplaceOne, onReplaceAll }) {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [hits, setHits] = useState([]);
  const [at, setAt] = useState(0);
  const [note, setNote] = useState(null);
  const [showReplace, setShowReplace] = useState(mode === 'replace');
  const atRef = useRef(0);
  atRef.current = at;
  const seq = useRef(0);

  const runSearch = useCallback(
    async (text, mc, { keepAt = false, note = null } = {}) => {
      const mine = ++seq.current;
      const found = text ? await onSearch(text, mc) : [];
      if (mine !== seq.current) return; // a later search landed first
      setHits(found);
      // A note the caller already has (Replace All's count) outranks the
      // search's own, or the count flashes and is gone.
      setNote(note || (text && !found.length ? 'No matches.' : null));
      const nextAt = found.length ? (keepAt ? Math.min(atRef.current, found.length - 1) : 0) : 0;
      setAt(nextAt);
      onGoto(found.length ? found[nextAt] : null);
    },
    [onSearch, onGoto]
  );

  useEffect(() => {
    const t = setTimeout(() => runSearch(find, matchCase), 150);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [find, matchCase]);

  const step = (dir) => {
    if (!hits.length) return;
    const next = (at + dir + hits.length) % hits.length;
    setAt(next);
    onGoto(hits[next]);
  };

  const replaceOne = async () => {
    if (!hits.length) return;
    const ok = await onReplaceOne(hits[at], replace);
    if (ok) await runSearch(find, matchCase, { keepAt: true });
  };

  const replaceAll = async () => {
    if (!find) return;
    const n = await onReplaceAll(find, replace, matchCase);
    await runSearch(find, matchCase, { note: `Replaced ${n} across the deck.` });
  };

  return (
    <div
      className="sl-find"
      // The pane sits inside the stage so it can be pinned over it; without
      // this the stage's own keyboard handler (nudge, delete, copy) would
      // see every keystroke typed here too, since it listens on a parent.
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === 'Escape') { onClose(); return; }
        if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
      }}
    >
      <div className="sl-find-row">
        <Icon name="find" size={14} />
        <input autoFocus className="rw-input sl-find-text" placeholder="Find" value={find} onChange={(e) => setFind(e.target.value)} />
        <span className="sl-find-count">{find ? `${hits.length ? at + 1 : 0} of ${hits.length}` : ''}</span>
      </div>
      <div className="sl-find-row">
        <Button icon="chevronUp" title="Previous match" className="sl-find-prev" disabled={!hits.length} onClick={() => step(-1)} />
        <Button icon="chevronDown" title="Next match" className="sl-find-next" disabled={!hits.length} onClick={() => step(1)} />
        <label className="sl-find-case">
          <input type="checkbox" className="sl-find-case-box" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          Match case
        </label>
        <Spacer />
        {showReplace ? null : <Button label="Replace…" className="sl-find-toggle-replace" onClick={() => setShowReplace(true)} />}
        <Button icon="close" title="Close" className="sl-find-close" onClick={onClose} />
      </div>
      {showReplace ? (
        <>
          <div className="sl-find-row">
            <input className="rw-input sl-find-with" placeholder="Replace with" value={replace} onChange={(e) => setReplace(e.target.value)} />
          </div>
          <div className="sl-find-row">
            <Button label="Replace" className="sl-find-replace-one" disabled={!hits.length} onClick={replaceOne} />
            <Button primary label="Replace all" className="sl-find-replace" disabled={!find} onClick={replaceAll} />
          </div>
        </>
      ) : null}
      {note ? <div className="sl-find-note">{note}</div> : null}
    </div>
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