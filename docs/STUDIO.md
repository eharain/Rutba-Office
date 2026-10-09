# Studio: the video editor, the advanced image editor and the recorder

The plan for one worker, building on `@rutba/studio` (see
[packages/studio/README.md](../packages/studio/README.md)). It sits beside
[ROADMAP.md](ROADMAP.md) and turns the Image and Video tiers of
[COMPETITORS.md](COMPETITORS.md) into work items on the code that is now
here. Read GAPS.md section 7 for what the two apps do today.

## 1. What this is for

Three things, all on one engine:

1. **Video** grows from trim-and-split into a full editor: several files on
   a timeline, tracks, titles, transitions, sound, keyframes, picture in
   picture, captions, and MP4 out.
2. **Image** gains an advanced editor as an extension of the app it is: the
   quick pipeline stays for crop, rotate and sliders, and an Advanced mode
   opens the same picture on the studio layer stack (layers, text, word
   art, shapes, adjustment layers, masks, blend modes).
3. **A recorder**: the screen, a window or a region, with the camera at the
   same time, the microphone and the computer's own sound, recorded together
   and landing on Video's timeline with the screen and the camera as two
   layers that stay editable. Presentations offers it for narration.

**It must run in a plain browser page as well as in the desktop app.**
Nothing in the editor's windows may reach Node, the file system or Electron
except through `shell` (the renderer's facade, `@rutba/office-shell`
client) and the engine's `configureMediaFetch` seam. Every engine and every
piece of document logic stays in a package with no DOM framework coupling,
so a web host can mount the same windows later with its own `shell`.

## 2. What came in

From the consumer suite's Studio at consumer `1aebcb67` (2026-10-09). Both
codebases are the same owner's; licensing and packaging are to be tidied
later.

| Here | From | Lines | State |
|---|---|---|---|
| `packages/studio/src/index.js` and the toolset files | `studio/packages/editor-core/*.js` | about 6,000 | Working, tested |
| `packages/studio/src/lib/` | `editor-core/lib/` | about 320 | Working |
| `packages/studio/src/doc.js`, `zoom.js`, `chart-data.js`, `insert-catalog.js` | `studio/apps/studio/lib/` | about 740 | Working; `insert-catalog` now imports `./host.js` |
| `packages/studio/src/host.js` | new, after `apps/studio/lib/renderer.js` | 70 | The window's one import; no transport |
| `packages/studio/src/capture/` | `consumer/packages/ui/lib/media-encode.js`, `audio-edit.js` | about 410 | Working, untested here |
| `packages/studio/port/` | Studio's editor components and the shared recorder and timeline | about 5,700 | React for Next and Bootstrap; not imported; to be ported |
| `tests/studio-*.test.js` | `studio/tests/*.test.mjs` | 10 files, 173 tests | Pass under `npm test` |

Left in consumer, on purpose: the Studio API, its database, render queue and
media client; the Next app's pages, uploads, shares, libraries, templates
and brand kits; decks and PPTX (Office has Presentations); the frozen A/B
harness (`editor-core/harness`, which compares frames against a byte-frozen
copy of the old painter). Consumer keeps working on its own copies until it
no longer needs them.

## 3. How the engine works, in one page

- A **document** (`src/doc.js`, `DOC_VERSION = 1`) is
  `{ v, kind, size, options, title, body, images, patches, at }`. `patches`
  are layer patches in the renderer's own format, so the document is the
  recipe with no translation step. Undo is a snapshot of it and autosave a
  write of it.
- `buildPlan` makes a plan from the document's options and media;
  `applyLayerPatches` puts the patches on it; `compileLayers` resolves
  fractions to pixels and orders the stack; `paintFrame(ctx, plan, t)`
  draws the instant `t`. A painter per layer type; unknown types are
  skipped so a newer document degrades rather than fails.
- **Time** lives on each layer: `timing { start, end }`, enter and exit
  ramps, and `keys` on fx, fy, fw, sizeFrac, opacity, rot, zoom, panX,
  panY. There is **one lane per layer** and no clip or track model; a
  `video` layer has a source `offset`, and its sound is a separate `sound`
  layer, so picture and sound trim apart.
- **Sound**: `soundLayers`, `clipFromSoundLayer` and `startAudioPreview`
  mix sound layers (volume, fades, loop, `mix | duck | solo`) through one
  `AudioContext`.
- **Output**: `renderVideo({ canvas, plan, audio, onProgress, signal })`
  plays the plan in real time into `captureStream(0)` and `MediaRecorder`
  with the mixed sound; `pickMimeType` prefers MP4 H.264 and falls back to
  WebM. Stills go through `image.js`.

## 4. What Office has today, and how it meets the engine

- **Video** (`apps/desktop/renderer/apps/video.js`, 462 lines) plays one
  file, splits (S), trims (I, O), removes a clip and exports WebM by
  recording a canvas. **Its export has no sound.** Its model is
  `@rutba/media`'s `Timeline` (`packages/media/src/timeline.js`): one track
  of `{ id, sourceId, start, end, speed, volume }` clips with split, trim,
  move and layout. `packages/media/package.json` exports `./export`, which
  does not exist.
- **Image** (`renderer/apps/image.js`, 428 lines) keeps an op list over
  `@rutba/imaging`'s pipeline (crop, rotate, flip, resize, adjust,
  annotate) and draws with `ctx.filter`.
- **Recording pieces already in Presentations**, all in
  `renderer/apps/slides/`:
  - `screen-record.js`: desktop capture through `getUserMedia` with a
    `chromeMediaSourceId` from `shell.capture.sources()` (the asking window
    left out), `MediaRecorder.start(500)` streamed to a scratch file with
    `shell.fs.temp`, `append` (4 MB slices) and `dropTemp`. No sound.
  - `cameo.js`: a shared, reference-counted camera stream and the patient
    open (`NotReadableError`, `AbortError`, `TrackStartError` retried with
    backoff for about nine seconds; a track that ends is asked for again).
  - `record.js` and `downsample.js`: microphone narration to WAV.
  - `export-video.js`: a canvas and mixed sound to `MediaRecorder`.
- `renderer/screenshot.js`: the source picker for screens and windows.
- **Main process**: `office-shell/src/electron/ipc.js` serves
  `capture.sources` and `capture.grab`; there is no
  `setDisplayMediaRequestHandler` and no permission handler yet.

## 5. Decisions to take first

Write each down in this file when taken.

- **D1. One model.** Recommended: the studio document is the project for
  both editors. A clip is a `video` layer with `timing` and `offset`; add a
  `lane` number to the layer envelope so the timeline draws tracks, and an
  optional `group` so a clip's picture and its sound move together.
  `@rutba/media`'s `Timeline` keeps its arithmetic (split, ripple, layout)
  as functions over those layers, or is retired once the editor no longer
  calls it. Do not keep two models.
- **D2. Project files.** `.rvid` and `.rimg`, each a zip of `doc.json`
  (the document), references to sources by path, and optionally the
  sources themselves ("collect files"). Autosave as the documents do.
  Register both in `packages/office-formats/src/registry.js`.
- **D3. Export.** Keep `renderVideo` (real time) as the first export, with
  sound. Then an encoder of our own: WebCodecs `VideoEncoder` and
  `AudioEncoder` into an MP4 muxer written in this repository (H.264 and
  AAC), frames painted as fast as the encoder takes them. It runs in a Web
  Worker, which needs an entry in `apps/desktop/build/bundle.js`; it fills
  `packages/media`'s missing `./export`.
- **D4. Media transport.** In the desktop app, `configureMediaFetch` reads
  through `rutba://file/` (byte ranges are served) and answers a Blob. A
  web host gives its own. No window fetches a path itself.
- **D5. Where UI lives.** Windows in `apps/desktop/renderer/apps/video/`
  and `apps/desktop/renderer/apps/image/` (folders, as `slides/` is), so
  `tools/find-english.mjs` and the message catalogue cover them. Engine
  logic that a web host would need goes in `packages/studio`.
- **D6. The recorder's file.** Recommended: two files recorded at once,
  the screen and the camera, each with its own `MediaRecorder` started on
  the same clock, placed as two layers (the camera a picture-in-picture
  layer with a mask), so the inset can be moved, resized or removed after.
  A single composited file is an export choice, not the recording.
- **D7. Consumer's second painter.** `consumer/packages/video/index.js` is
  a diverged copy of the same renderer. From now on Office's copy is the
  one developed; nothing is synced back by hand.

## 6. The work, in order

Each phase is releasable on its own and follows docs/RELEASING.md (a
feature takes the minor number, the note names its checks).

### Phase 1. Video on the studio engine (Video Tier 1)

1. Port `useCreative`, `Stage` and `ui/VideoTimeline` into
   `renderer/apps/video/`, on `office-ui`, every word through `t()` with
   Urdu. The stage paints `paintFrame` at the playhead; play and scrub
   drive `t`.
2. Open several files: each video becomes a `video` layer and a `sound`
   layer, end to end on lane 1. Stills become `image` layers with a
   duration. Pictures' "In Video" and Presentations' Export to Video open
   here.
3. **Split** (new; the studio has none): a layer split at `t` becomes two
   layers, the second with `offset` moved on by the same amount, its
   sound layer split with it. Trim by dragging either end, ripple delete,
   move, speed, volume, mute, fade in and out.
4. Titles and lower thirds (`title`, `text`, word art), transitions through
   the enter and exit ramps, crop and rotate per clip.
5. Export with sound through `renderVideo`; the existing WebM export is
   replaced. Undo through `@rutba/editing/history`.
6. `.rvid` (D2), autosave and recovery as the documents have.

Done when: window checks open two clips and a still, split, trim, add a
title and a crossfade, export, and read the exported file's duration, a
frame's pixels and that it has an audio track; engine tests for split and
ripple; the whole gate green.

### Phase 2. Tracks, keyframes and sound (Video Tier 2)

1. Lanes (D1) drawn as tracks; picture in picture with a `video` layer's
   fx, fy, fw and mask; overlays.
2. A keyframe row per layer for position, size, opacity, rotation and
   volume, on the engine's existing `keys`, with ease choices.
3. Audio tracks, voice-over recorded in the app (`capture/audio-edit.js`,
   Presentations' narration recorder), music ducked under speech (`mix:
   'duck'` exists), beat snapping (`lib/beats.js`).
4. Captions: SRT and WebVTT in and out, edited in a list, burnt in (a
   `caption` layer) or kept as a track.
5. Chroma key and `.cube` LUTs as new adjustment layer types registered
   through `registerLayerType`; colour (lift, gamma, gain, temperature).

### Phase 3. The recorder

1. A Record command in Video (and in Presentations, Insert → Screen
   Recording, which today uses `screen-record.js`).
2. Sources: a screen, a window, or a region of a screen (recorded whole and
   cropped by the layer, so the crop can change); the camera
   (`cameo.js`'s patient open, moved to a shared module); the microphone;
   the computer's sound. The last needs the main process to answer
   `getDisplayMedia` with `session.setDisplayMediaRequestHandler`
   (`audio: 'loopback'` on Windows); add it in `office-shell` with a
   contract entry, as every capability is.
3. Two recorders on one clock (D6), each streamed to its own scratch file
   (`shell.fs.temp`, `append`, `dropTemp`, as `screen-record.js` does), a
   live preview with the camera inset where it will sit, a countdown,
   pause and resume, and a hotkey to stop. The camera can be left out, or
   recorded alone.
4. On stop the takes land on the timeline: the screen on lane 1, the camera
   above it as an inset with a rounded mask, the sound as its own layer.
5. Checks run with `RUTBA_FAKE_MEDIA=1` (Chromium's stand-in camera and
   microphone) and a fake screen source; the stand-in camera stays stopped
   once a recording stops it, so recorder checks run last, as the cameo
   checks do.

### Phase 4. The advanced image editor

1. Image's ribbon gains Advanced (and Edit in Advanced from the right-click
   menu): the picture opens as a one-frame studio document, the quick
   pipeline's crop, rotate and sliders carried over as its first layer's
   crop and an adjustment layer.
2. Layers panel (order, visibility, opacity, blend mode, mask), text, word
   art, shapes, icons, adjustment layers from the toolset; export through
   `renderImage` at the source's size.
3. `.rimg` (D2). Save back to the picture's own format replaces nothing
   until asked, as Image does now.
4. New engine work, not in the studio: rectangle, ellipse, lasso and magic
   wand selections with feather; a brush, a clone stamp and healing;
   perspective. Each a module in `packages/studio` with its own tests.
5. Picture Format → Edit in Image from Documents, Worksheets and
   Presentations opens Advanced and puts the result back.

### Phase 5. Faster export (Video Tier 3)

The WebCodecs encoder and MP4 muxer (D3), faster than real time, with 720p,
1080p and 4K presets and a bitrate choice; WebM stays as the free-codec
option; GIF and audio-only; a frame to Image; proxies for 4K on a laptop.

### Phase 6. Ready for a web host

No window in Video or Image touches anything outside `shell` and the
engine; a check that runs the engine tests in a browser page; a list in
this file of what a web `shell` must provide (`fs.read`, `fs.write`,
`fs.temp`, `append`, `dropTemp`, `capture.sources` or the browser's own
`getDisplayMedia`, the media fetch).

## 7. Conventions that bite

- **A new workspace package** is declared in `apps/desktop/package.json`
  or the installer misses it (`tests/desktop-deps.test.js`); `@rutba/studio`
  already is.
- **Tests** go in the root `tests/` as `*.test.js`; tests inside a package
  are not found.
- **Words**: every visible word through `t()`, with Urdu in
  `catalogues/ur.js`; `npm run` the message extract; `tests/messages.test.js`,
  `catalogue-ur.test.js` and `english-left.test.js` must pass. An engine
  error a user sees goes in `renderer/engine-words.js`.
- **Window checks** live in `apps/desktop/main/verify-*.js` and run under
  `npm run gate`; see docs/TESTING.md.
- **The toolset must be registered** before painting: import
  `@rutba/studio/host`. Keep `sideEffects` listing every file that
  registers a layer type.
- **`loadImages`** from `host` answers an array; the engine's own answers a
  report object, and passing that to `buildPlan` draws nothing, silently.
- **Media as Blobs**, never a cross-origin URL, or the canvas taints and
  `captureStream` and adjustment layers stop working.
- **Two tests read `port/` files as text** (`studio-insert`, `studio-zoom`);
  move them with the component when it is ported.
- **Commits**: the repository's hooks refuse names of AI tools and providers
  in messages, file names and file content.

## 8. Open questions for the owner

- Should the recorder's default be the two-file recording (D6) or one
  composited file?
- Is Advanced a mode inside Image, or a separate app on the launcher?
- Do Image and Video keep their names when they become editors?
- Which comes first after Phase 1: the recorder (Phase 3) or tracks
  (Phase 2)?
