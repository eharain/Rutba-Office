# port/: the Studio editor's UI, to be rewritten

These files are the React UI of the consumer suite's Studio editor and the
shared recorder and timeline components, copied as they were. **Nothing
imports them and the bundler never sees them.** They are here so the Video
and Image editors can be built from working code rather than from memory;
each is deleted once its Office version lands.

They use Next.js, Bootstrap and Font Awesome class names, `@rutba/ui`
(`ProtectedRoute`, `AuthContext`, the recorder and timeline components),
`@rutba/api-client` and the Studio API (`lib/studio-api`). None of that
exists in Office. A port keeps the behaviour and replaces the frame:

- `@rutba/office-ui` for the ribbon, commands (`useCommands`), panels and
  dialogs;
- `t()` and `tn()` for every word a window shows, with an Urdu entry in
  `catalogues/ur.js` (`tests/catalogues.test.js` refuses a missing one);
- `shell.fs` and the `rutba://file/` protocol instead of the media proxy,
  uploads and libraries;
- `@rutba/studio/host` instead of `../../lib/renderer`.

| File | What it does there | Becomes |
|---|---|---|
| `editor/useCreative.js` | The document, selection, undo (`@rutba/editing/history`, identical in Office), autosave | The editor's state hook |
| `editor/Stage.js` | The canvas: paint at the playhead, hit test, drag, resize, zoom and pan | The editor's stage |
| `editor/LayerStack.js` | The layer list with z order and visibility | A pane |
| `editor/Inspector.js` | Every layer type's properties (1,033 lines) | The properties pane, split by type |
| `editor/controls.js` | Sliders, colour, number and select inputs | `office-ui` controls |
| `editor/DocSettings.js` | Aspect, theme, fit, size | A ribbon group or pane |
| `editor/EditorPanes.js`, `LeftRail.js` | The editor's layout | `AppFrame` and panes |
| `editor/InsertPanel.js` | The insert panel over `insert-catalog` | The Insert tab |
| `editor/Timeline.js` | The adapter onto `ui/VideoTimeline` | The timeline pane |
| `editor/MusicBedPanel.js`, `AudioTracks.js` | Music bed and sound layers (library fetch is consumer only) | Audio track controls |
| `editor/CreativeEditor.js` | The page that puts it together, with shares, templates and uploads | Reference only |
| `ui/VideoTimeline.js` | Lanes per layer: move, trim, reorder | The timeline |
| `ui/RecorderDialog.js` | Camera or screen video; microphone, system sound or both mixed | The recorder's base |
| `ui/VideoEditorDialog.js` | Trimming a take by re-recording it | Reference for trim on record |
| `ui/RecordPicker.js` | Choosing what to record | The recorder's source picker |
| `renderer.js` | The consumer page's seam: media proxy with a token | Done as `src/host.js` |
