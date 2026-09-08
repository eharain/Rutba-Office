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
import { Ribbon, Group, Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Dialog, Field, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import Presenter from './slides/presenter.js';

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
  const stageRef = useRef(null);
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

  useEffect(() => {
    const template = new URLSearchParams(location.search).get('template');
    const run = async () => {
      setBusy(true);
      try {
        const opened = boot.file
          ? await shell.doc.open({ path: boot.file, width: 1280 })
          : await shell.doc.new({ kind: 'slides', template: template && template !== 'blank' ? template : 'deck' });
        setDoc(opened);
        setModel(opened.model);
        if (opened.path) shell.app.addRecent({ path: opened.path, app: 'slides' }).catch(() => {});
        if (opened.converted?.from) toast(`Opened from ${opened.converted.from.toUpperCase()}. Saving will write a .pptx.`, { ms: 5200 });
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
      'view.present': { label: 'Present', icon: 'play', key: 'F5', run: () => setPresent(true) },
      'slide.add': { label: 'New slide', icon: 'plus', key: 'Mod+M', run: () => addSlideRef.current?.('obj') },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, model, index, apply, save, openFile, shell]
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
    shell.win.fullscreen({ on: true });
    return () => {
      window.removeEventListener('keydown', onKey);
      shell.win.fullscreen({ on: false });
    };
  }, [present, model, shell]);

  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Presentation" menu={appMenu}>
        <Empty icon="slides" title="This file could not be opened">{error}</Empty>
      </AppFrame>
    );
  }

  const slide = model?.slide;

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
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'insert', label: 'Insert' },
            { id: 'design', label: 'Design' },
            { id: 'show', label: 'Slide Show' },
            { id: 'view', label: 'View' },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="save" title="Save" onClick={() => save(false)} />
              <Button icon="undo" title="Undo" disabled={!doc?.canUndo} onClick={() => commands['edit.undo']?.run?.()} />
              <Button icon="play" title="Present" onClick={() => setPresent(true)} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="Slides">
                <Button
                  tall
                  icon="plus"
                  label="New slide"
                  onClick={(e) =>
                    menu.open(e, [
                      { label: 'Title and content', icon: 'slides', run: () => addSlide('obj') },
                      { label: 'Title slide', icon: 'slides', run: () => addSlide('title') },
                      { label: 'Blank', icon: 'file', run: () => addSlide('blank') },
                    ])
                  }
                />
                <Button icon="copy" label="Duplicate" onClick={() => commands['slide.new'].run()} />
                <Button icon="trash" label="Delete" onClick={() => commands['slide.delete'].run()} disabled={(model?.count || 0) < 2} />
              </Group>
              <Group label="Arrange">
                <Button icon="chevronUp" label="Move up" disabled={index === 0} onClick={() => { apply({ op: 'moveSlide', from: index, to: index - 1 }); setIndex(index - 1); }} />
                <Button icon="chevronDown" label="Move down" disabled={index >= (model?.count || 1) - 1} onClick={() => { apply({ op: 'moveSlide', from: index, to: index + 1 }); setIndex(index + 1); }} />
              </Group>
              <Group label="Notes">
                <Button tall icon="word" label="Speaker notes" onClick={() => setNotesOpen(true)} />
              </Group>
              <Group label="File">
                <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'slides' })} />
                <Button tall icon="open" label="Open" onClick={openFile} />
                <Button tall icon="save" label="Save" onClick={() => save(false)} />
              </Group>
            </>
          ) : tab === 'insert' ? (
            <>
              <Group label="Text">
                <Button tall icon="textbox" label="Text box" onClick={() => commands['slide.textbox'].run()} />
              </Group>
              {/*
                The deck engine writes text boxes and nothing else yet. A Shape
                button that quietly inserted a text box was here for a day; a
                control that does something other than what it says is worse
                than one that is missing, and the audit says shapes are missing.
              */}
              <Group label="Slides">
                <Button tall icon="plus" label="New slide" onClick={() => addSlide('obj')} />
                <Button tall icon="copy" label="Duplicate" onClick={() => commands['slide.new'].run()} />
              </Group>
            </>
          ) : tab === 'design' ? (
            <>
              <Group label="This deck">
                <Button
                  tall
                  icon="grid"
                  label={model?.size ? `${Math.round(model.size.width)} × ${Math.round(model.size.height)}` : 'Slide size'}
                  title={model?.size ? `Slides are ${Math.round(model.size.width)} by ${Math.round(model.size.height)} pixels — ${Math.abs(model.size.width / model.size.height - 16 / 9) < 0.02 ? '16:9 widescreen' : Math.abs(model.size.width / model.size.height - 4 / 3) < 0.02 ? '4:3 standard' : 'a custom ratio'}. Colours and fonts come from the deck's own theme.` : 'The deck decides its own size and theme'}
                  disabled
                />
              </Group>
              <Group label="Layout">
                <Button tall icon="slides" label="Title slide" onClick={() => addSlide('title')} />
                <Button tall icon="slides" label="Title and content" onClick={() => addSlide('obj')} />
                <Button tall icon="file" label="Blank" onClick={() => addSlide('blank')} />
              </Group>
              <Group label="Export">
                <Button tall icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />
              </Group>
            </>
          ) : tab === 'show' ? (
            <>
              <Group label="Start">
                <Button tall icon="play" label="From the start" onClick={() => { setIndex(0); setPresent(true); }} />
                <Button tall icon="play" label="From here" onClick={() => setPresent(true)} />
                <Button tall icon="grid" label="Presenter view" title="Opens a second window with your notes, the next slide and a clock — put it on the other screen" onClick={presentWithNotes} />
              </Group>
              <Group label="Notes">
                <Button tall icon="word" label="Speaker notes" onClick={() => setNotesOpen(true)} />
              </Group>
            </>
          ) : (
            <>
              <Group label="Window">
                <Button tall icon="maximize" label="Full screen" onClick={() => shell.win.fullscreen({})} />
              </Group>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
                <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
                <Button icon="check" label="100%" onClick={() => shell.win.zoom({ reset: true })} />
              </Group>
            </>
          )}
        </Ribbon>
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
                  <span className="sl-thumb-card">
                    <span className="sl-thumb-title">{o.title || 'Untitled slide'}</span>
                  </span>
                </button>
              ))}
            </div>
          </Panel>

          <Content>
            <div className="sl-stage" ref={stageRef}>
              {slide ? (
                <div className="sl-slide" style={{ width: model.size.width, height: model.size.height }}>
                  <div className="sl-svg" dangerouslySetInnerHTML={{ __html: slide.svg }} />
                  {/* Text boxes get a hit area so a click lands on the shape rather than on the drawing. */}
                  {slide.shapes.filter((s) => s.text && s.geometry).map((s) => (
                    <button
                      key={s.id}
                      type="button"
                      className="sl-hit"
                      style={{ left: s.geometry.x, top: s.geometry.y, width: s.geometry.w, height: s.geometry.h }}
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
              ) : (
                <Empty icon="slides" title="This presentation has no slides" />
              )}
            </div>
            {slide?.notes ? <div className="sl-notes">{slide.notes}</div> : null}
          </Content>
          {menu.node}
        </>
      )}

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

.sl-stage { flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center; padding: 22px; background: var(--window); }
.sl-slide { position: relative; box-shadow: var(--shadow-2); background: #fff; }
.sl-svg svg { display: block; width: 100%; height: 100%; }
.sl-hit { position: absolute; border: 1px solid transparent; background: transparent; border-radius: 2px; }
.sl-hit:hover { border-color: var(--accent-line); background: rgba(43, 95, 217, 0.06); }
.sl-editor {
  position: absolute; border: 2px solid var(--accent); border-radius: 3px; padding: 4px 6px;
  font: inherit; font-size: 15px; background: #fff; color: #111; resize: none; outline: none; z-index: 5;
}
.sl-notes {
  border-top: 1px solid var(--line); background: var(--chrome); padding: 9px 18px;
  font-size: 12.5px; color: var(--ink-2); max-height: 110px; overflow: auto; white-space: pre-wrap;
}

.sl-present { position: fixed; inset: 0; background: #000; display: grid; place-items: center; z-index: 200; }
.sl-present-stage { width: min(100vw, 177.78vh); }
.sl-present-stage svg { display: block; width: 100%; height: auto; }
.sl-present-bar {
  position: fixed; bottom: 14px; left: 50%; transform: translateX(-50%);
  color: rgba(255,255,255,0.55); font-size: 12px; letter-spacing: 0.02em;
}
`;

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