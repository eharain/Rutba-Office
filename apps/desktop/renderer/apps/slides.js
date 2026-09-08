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
import { Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Dialog, Field, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
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
  const [present, setPresent] = useState(false);
  const [notesOpen, setNotesOpen] = useState(false);
  const [blank, setBlank] = useState(false);
  // Set when a presenter window is driving, so this one follows rather than leads.
  const [led, setLed] = useState(false);
  /**
   * How the deck is shown. None of it is in the file: the view mode, the
   * rulers, gridlines and guides, whether the notes strip shows, a zoom
   * level (null fits the window), and the colour/greyscale tone.
   */
  const [view, setView] = useState({ mode: 'normal', ruler: false, gridlines: false, guides: false, notes: true, zoom: null, tone: 'colour' });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  // The selected shape — one click selects, a double-click edits its words —
  // is what the Font and Paragraph groups act on.
  const [selected, setSelected] = useState(null);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  // Reading View: the show in this window, without going full screen.
  const [reading, setReading] = useState(false);
  const stageRef = useRef(null);
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
        const opened = boot.file
          ? await shell.doc.open({ path: boot.file, kind: 'deck', width: 1280 })
          : await shell.doc.new({ kind: 'slides', template: template && template !== 'blank' ? template : 'deck' });
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

  const commitText = useCallback(
    async (shapeId, text) => {
      const paragraphs = String(text).split('\n').map((line) => ({ runs: [{ text: line }] }));
      await apply({ op: 'setText', slide: index, shape: shapeId, paragraphs });
      setEditing(null);
    },
    [apply, index]
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
    async (format) => {
      if (!doc) return;
      const target = await shell.dialog.save({
        title: `Export as ${format.toUpperCase()}`,
        defaultPath: (doc.path || doc.name).replace(/\.[^.]+$/, `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format] }],
      });
      if (!target) return;
      try {
        await shell.doc.export({ id: doc.id, format, path: target });
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
  useEffect(() => { setSelected(null); }, [index]);


  const slide = model?.slide;
  const selectedShape = selected ? slide?.shapes?.find((s) => s.id === selected) || null : null;
  // What the ribbon shows for the selected shape: its first run's look and
  // its first paragraph's alignment — the granularity the writer edits at.
  const format = useMemo(() => {
    const p = selectedShape?.text?.paragraphs?.[0];
    const r = p?.runs?.[0] || {};
    return { bold: Boolean(r.bold), italic: Boolean(r.italic), underline: Boolean(r.underline), size: r.size || 18, color: r.color || null, font: r.font || '', align: p?.align || 'left' };
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
        const paragraphs = selectedShape.text.paragraphs.map((p) => {
          const { plain, ...rest } = p;
          return {
            ...rest,
            align: arg.align ?? p.align,
            runs: (p.runs || []).map((r) => ({
              ...r,
              bold: arg.bold === 'toggle' ? !r.bold : arg.bold ?? r.bold,
              italic: arg.italic === 'toggle' ? !r.italic : arg.italic ?? r.italic,
              underline: arg.underline === 'toggle' ? !r.underline : arg.underline ?? r.underline,
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
                <button
                  key={o.part || i}
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
              ))}
            </div>
          </Panel>

          <Content>
            <div
              className={`sl-stage tone-${view.tone || 'colour'}`}
              ref={stageRef}
              onMouseDown={(e) => { if (e.target === e.currentTarget) setSelected(null); }}
            >
              {slide && view.mode === 'sorter' ? (
                <div className="sl-sortergrid">
                  {(model.outline || []).map((o, i) => (
                    <button key={o.part || i} type="button" className={`sl-sortercard${i === index ? ' active' : ''}`} onClick={() => { setIndex(i); patchView({ mode: 'normal' }); }} title={o.title || `Slide ${i + 1}`}>
                      {o.thumbnail ? <span className="sl-thumb-pic" dangerouslySetInnerHTML={{ __html: o.thumbnail }} /> : <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>}
                      <span className="sl-sortern">{i + 1}</span>
                    </button>
                  ))}
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
                  {/* Text boxes get a hit area so a click lands on the shape rather than on the drawing. */}
                  {slide.shapes.filter((s) => s.text && s.geometry).map((s) => (
                    <button
                      key={s.id}
                      type="button"
                      className={`sl-hit${selected === s.id ? ' selected' : ''}`}
                      style={{ left: s.geometry.x, top: s.geometry.y, width: s.geometry.w, height: s.geometry.h }}
                      onClick={() => setSelected(s.id)}
                      onDoubleClick={() => setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') })}
                      onContextMenu={(e) => menu.open(e, [
                        { label: 'Edit text', icon: 'textbox', run: () => setEditing({ id: s.id, text: s.text.paragraphs.map((p) => p.plain).join('\n') }) },
                        { label: 'Delete shape', icon: 'trash', run: () => apply({ op: 'removeShape', slide: index, shape: s.id }) },
                      ])}
                      title={`${s.name || 'Shape'} — double-click to edit`}
                    />
                  ))}
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
          {menu.node}
        </>
      )}

      {shortcutsOpen ? <SlidesShortcutsDialog onClose={() => setShortcutsOpen(false)} /> : null}

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

const CSS = `
.sl-sorter { display: flex; flex-direction: column; gap: 6px; padding: 8px; }
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
.sl-hit { position: absolute; border: 1px solid transparent; background: transparent; border-radius: 2px; min-height: 8px; min-width: 8px; }

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