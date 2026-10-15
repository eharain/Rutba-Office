# Studio: the video editor, the advanced image editor and the recorder

The plan for one worker, building on `@rutba/studio` (see
[packages/studio/README.md](../packages/studio/README.md)). It sits beside
[ROADMAP.md](ROADMAP.md) and turns the Image and Video tiers named there
into work items on the code that is now here. Read GAPS.md section 7 for what the two apps do today.

## 1. What this is for

Five things, all on one engine, each set out in its own section below:

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
4. **Introductory and explainer videos**: a guided flow that records a
   person's screen and camera scene by scene, with a script on a
   teleprompter, layouts that can be switched after recording, a brand kit,
   captions and music, and exports for each destination (section 9).
5. **Performance and formats**: frame-exact decoding, GPU compositing,
   proxies and export faster than real time, held to budgets as tests
   (section 10); and the formats that open and the ones written, video,
   sound and pictures, each with its settings (section 11).

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
  `capture.sources` and `capture.grab`, and since Phase 2 `choose`
  (the `setDisplayMediaRequestHandler`), `controls`, `space` and
  `take`; there is no permission handler yet.

## 5. Decisions to take first

Write each down in this file when taken.

- **D1. One model.** Recommended: the studio document is the project for
  both editors. A clip is a `video` layer with `timing` and `offset`; add a
  `lane` number to the layer envelope so the timeline draws tracks, and an
  optional `group` so a clip's picture and its sound move together.
  `@rutba/media`'s `Timeline` keeps its arithmetic (split, ripple, layout)
  as functions over those layers, or is retired once the editor no longer
  calls it. Do not keep two models.

  **Taken (Phase 1, 2026-10-12), as recommended.** The project is a studio
  document (`kind: 'video'`, `options.bare`: the plan compiles none of the
  slideshow's own layers and runs to the end of its last clip). A video clip
  is a `video` layer with `timing` and `offset`, its sound a `sound`
  layer on the same source in the same `group`; a still is an `image`
  layer with `fit`; a title is a `text` layer, or a shape and a text in
  one group. The layer envelope gained `lane` (counted from 0 within the
  pictures and within the sound) and `group`, and whole pictures gained
  `fit` (`cover`, `contain`, `blur`), a quarter `turn` and flips. The
  timeline's arithmetic (add, split, trim, ripple, gaps, move, copy and paste,
  snapping, links, markers, transitions, titles) is
  `packages/studio/src/project.js`, pure functions over the document, tested
  in `tests/studio-project.test.js`; `planPatches` hands the renderer the
  pictures lane by lane, so a lane's order is its paint order. Two clips of
  one source that overlap (a crossfade between the halves of a split) are
  given a decoder each through a `videoKey`. The Video app no longer calls
  `@rutba/media`'s `Timeline`; it is left for Pictures and any other caller
  and is not developed further.
- **D2. Project files.** `.rvid` and `.rimg`, each a zip of `doc.json`
  (the document), references to sources by path, and optionally the
  sources themselves ("collect files"). Autosave as the documents do.
  Register both in `packages/office-formats/src/registry.js`.

  **Taken for `.rvid` (Phase 1, 2026-10-12).** `packages/studio/src/rvid.js`
  writes a stored zip: `mimetype` first
  (`application/vnd.rutba.video-project+zip`, so the sniffer knows it without
  its name), `doc.json` and `sources.json` (each source's path, name, size,
  length, kind and its path relative to the project). On open a source missing
  at its path is looked for at its relative path, and the clips are pointed
  there; what is still missing is named. Collecting the files into the zip is
  V35. Autosave writes the same bytes to the documents' recovery folder
  (`doc.keepCopy`), and the launcher offers them back.

  **Taken for `.rimg` (Image advanced editor, part 1, 2026-10-12).**
  `packages/studio/src/rimg.js` writes a stored zip like `.rvid`: `mimetype`
  first (`application/vnd.rutba.image-project+zip`), `doc.json` (the image
  document: the canvas size in pixels and the layers, with the quick edits
  made to the opened picture under `quick`), `source/<name>` (the picture
  as it was opened, untouched) and `bitmaps/<name>` (the painted bitmaps:
  the opened picture with its quick edits, and every picture added as a
  layer). Unlike `.rvid` nothing is referenced outside the file; a layer
  names a bitmap inside it as `rimg:<path>`, and a missing one is named on
  open. On open the quick edits are made again from the source, so they stay
  editable. `.rimg` is registered in `packages/office-formats` (sniffed by
  its first entry, offered by Open and Save As, associated with Image), and
  autosave writes the same bytes through `doc.keepCopy` (kind `image`), which
  the launcher's Recover opens in Image.
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

  **Taken (Phase 2, 2026-10-12), as recommended.** The screen and the camera
  are two WebM files, each with its own `MediaRecorder`, started on the same
  tick and paused and resumed together. The microphone goes into the
  camera's file and the computer's own sound into the screen's, so each
  sound lands on a lane of its own; without a camera both go into the
  screen's file, mixed. A file with no sound to record carries a silent
  track (a picture-only recorder on a window that seldom changes kept the
  paused stretch in its file). On the timeline the camera is a `video`
  layer with `fx`, `fy`, `fw`, `fh`, an `insetShape` (`circle`,
  `rounded`, `rect`) and `flipX` for the mirror; the screen is a
  `video` layer with `fit: 'contain'` and, for a region, a `crop`. All
  of a recording's layers share one group. One composited file is what
  Export makes.
- **D7. Consumer's second painter.** `consumer/packages/video/index.js` is
  a diverged copy of the same renderer. From now on Office's copy is the
  one developed; nothing is synced back by hand.

## 6. The video editor

Every item names where it lives and what proves it. "Engine" means
`packages/studio` with tests in `tests/studio-*`; "window" means
`renderer/apps/video/` with a window check in `main/verify-*.js`.

### 6.1 Timeline and editing

| # | Item | Notes |
|---|---|---|
| V1 | Several files on one timeline, end to end | Each video becomes a `video` layer and a `sound` layer grouped (D1); stills become `image` layers with a duration (default 5 s, a setting) |
| V2 | Split at the playhead (S), and Split All Tracks (Shift+S) | New in the engine: a layer split at `t` becomes two, the second's `offset` moved on by the same amount, its grouped sound split with it |
| V3 | Trim by dragging either end; Trim Start (I) and Trim End (O) | Rolls the source `offset` on a start trim |
| V4 | Ripple delete, ripple trim, and a ripple switch | Off: a gap stays where a clip was removed; a Close Gap command closes it |
| V5 | Move, copy, paste, duplicate, and drag between lanes | Snapping to clip edges, the playhead, markers and beats (`lib/beats.js`); Alt drag skips snapping |
| V6 | Detach audio, and unlink and relink a group | A detached sound layer trims and moves on its own |
| V7 | Markers with names, and Go to Next and Previous Marker | Kept in the document; chapters on export (11.2) |
| V8 | Zoom the timeline (Ctrl+wheel, Fit), scroll, follow the playhead | Thumbnails per clip and a waveform per sound layer, cached (P2, P9) |
| V9 | Undo and redo of every edit | `@rutba/editing/history` snapshots of the document |
| V10 | Frame stepping (comma, period), J K L shuttle, In and Out points | The keys Premiere, DaVinci and Shotcut users expect |
| V11 | Freeze frame, reverse, and speed with pitch kept | Speed already in `@rutba/media`; pitch kept through the audio rate on export |
| V12 | Lanes as tracks, any number of video and audio tracks, mute, solo, lock, hide | D1's `lane`; a lane's order is its z order |

### 6.2 Picture

| # | Item | Notes |
|---|---|---|
| V13 | Crop, rotate, flip, fit, fill and letterbox per clip | Engine already has crop, zoom, pan and `fit: blur` |
| V14 | Picture in picture: position, size and a mask (rectangle, rounded, circle) | `video` layer fx, fy, fw and mask; a rounded rectangle is new |
| V15 | Transitions between clips: crossfade, dip to black or white, wipe, slide, push, zoom | Enter and exit ramps exist; a transition is the outgoing clip's exit and the incoming clip's enter over an overlap |
| V16 | Colour: exposure, contrast, saturation, temperature, tint, lift, gamma, gain, and `.cube` LUTs | Adjustment layers exist for most; lift, gamma, gain and LUTs are new layer types through `registerLayerType` |
| V17 | Effects: blur, sharpen, vignette, film grain, black and white, sepia | Mostly adjustment layers; grain is new |
| V18 | Chroma key (green screen) with spill suppression | New layer effect; a pixel pass in a worker or a WebGL shader (P3) |
| V19 | Ken Burns on stills (pan and zoom across a picture) | Keys on zoom and pan, with presets |
| V20 | Stabilisation | Tier 3: feature tracking, analysed once and cached |

### 6.3 Words on screen

| # | Item | Notes |
|---|---|---|
| V21 | Titles, lower thirds, end cards and callouts from a gallery | `title`, `text`, word art, shapes; a gallery of presets in the insert catalogue |
| V22 | Text animation in and out (fade, rise, type on, pop) | Ramps exist; type on is the `caption` typewriter |
| V23 | Captions: import and export SRT and WebVTT, edit in a list, style, burn in or keep as a track | New engine module for the formats; `caption` layers to draw them |
| V24 | Captions from speech on this computer | Tier 3: a local speech model, a one-time consented download, in the languages the suite is localised to; never a service |
| V25 | Right-to-left titles and captions (Urdu, Arabic) | The engine's text measures with the canvas; check Nastaliq shaping and line breaking |

### 6.4 Sound

| # | Item | Notes |
|---|---|---|
| V26 | Volume per clip and per track, mute, fade in and out | `sound` layer volume and fades exist |
| V27 | Volume keyframes drawn on the waveform | Engine keys gain `volume` |
| V28 | Music bed with ducking under speech | `mix: 'duck'` exists; detection of speech by level |
| V29 | Voice-over recorded straight onto a track | Presentations' narration recorder (`slides/record.js`) moved to a shared module |
| V30 | Noise reduction and loudness normalisation (to -14 LUFS for the web, -16 for speech) | New: an offline pass on export; normalisation needs a loudness meter (EBU R128) |
| V31 | Audio-only export | 11.3 |

### 6.5 Project and files

| # | Item | Notes |
|---|---|---|
| V32 | `.rvid` project: a zip of `doc.json` and source references | D2; registered in `office-formats` |
| V33 | Autosave and recovery after a crash | As the documents have: the launcher offers it back |
| V34 | Relink missing media | A source moved or renamed is asked for, matched by name and size |
| V35 | Collect files: a project and its media in one folder or one zip | For handing a project to someone |
| V36 | Open a whole folder of clips; drag files in from Pictures or Explorer | |

## 7. The image editor

Image keeps its quick editor for the everyday jobs and gains an Advanced
editor on the studio layer stack (Phase 5). Items marked Quick land in the
current app; Advanced items need the layer stack.

### 7.1 Quick edits (Image Tier 1)

| # | Item | Notes |
|---|---|---|
| I1 | Buttons for what the pipeline already does: flip vertical, resize, invert, and the annotation tools (rectangle, ellipse, arrow, text, pen) with colour and width | Quick; `@rutba/imaging` has them with no UI |
| I2 | Straighten by angle, with the crop following | Quick |
| I3 | Auto enhance, levels, curves, white balance, exposure, highlights and shadows | Quick; studio's `adjust.js` has the arithmetic for levels and curves |
| I4 | Sharpen, noise reduction, vignette, red-eye | Quick |
| I5 | Copy to the clipboard, and paste a picture in as a new image | Quick |
| I6 | Batch from Pictures: resize, convert, rename by pattern, rotate by EXIF, strip metadata | Quick; a dialog over a folder, run off the main process |

### 7.2 Advanced editor (Image Tier 2)

**Decided (part 1, 2026-10-12): Advanced is a mode inside Image, and the
app keeps the name Image.** A plain picture opens as it always did, with the
quick edits of 7.1. Adding a layer (anything from the Insert tab, Paste onto
a layered picture, Use layers, or Canvas Size) or opening a `.rimg` turns on
the Layers panel: the opened picture becomes the Background layer, its quick
edits still applied to it from the Home, Adjust, Filters and Draw tabs, and
the canvas is a studio image document (`packages/studio/src/layers.js`)
painted by the engine. With layers, Crop cuts the canvas and everything on
it; Rotate, Flip and Straighten still turn the Background picture; Resize
becomes Image Size (the whole picture). Save writes a `.rimg`; Export
flattens to PNG, JPEG (clear pixels white) or WebP through the same export
and its quality, size and longest-side options. Undo covers the layers and
the quick edits together.

The layer model is built for I10 to I14: a layer is a studio patch, a group
is a patch of its own (`type: 'group'`) whose members name it as `parent`
and sit beneath it in the list, and `reframe` (canvas size, trim, crop)
already carries a layer's mask fractions with it. A painted mask, a
selection, a brush stroke or a clone will be fields on a layer's patch (an
alpha bitmap kept under `bitmaps/` in the `.rimg`, as pictures are).

| # | Item | Notes |
|---|---|---|
| I7 | Layers panel: order, visibility, opacity, blend mode, lock, group | Studio's layer stack and `blend` |
| I8 | Text, word art, shapes, icons and charts as layers | Studio's toolset as it is |
| I9 | Adjustment layers that affect what is under them | `adjust.js` |
| I10 | Layer masks, painted and feathered | Engine masks are hard-edged shapes; a painted, feathered mask is new (an alpha bitmap per layer) |
| I11 | Selections: rectangle, ellipse, lasso, magic wand, select by colour, with feather, grow and invert | New engine module; every adjustment and fill applies inside the selection |
| I12 | Brush, eraser, fill and gradient | New; pressure from a pen where the platform gives it |
| I13 | Clone stamp and healing brush | New; healing by patch blending |
| I14 | Perspective correction and free transform with corner handles | New; `resizePatch` covers scale only |
| I15 | Canvas size, image size, and trim transparent edges | |
| I16 | Edit in Image from Documents, Worksheets and Presentations, saving back in place | Picture Format → Edit in Image |
| I17 | `.rimg` project: a zip of `doc.json`, the source and painted bitmaps | D2 |

#### Part 1 as built (2026-10-12)

I7, I8, I9, I15, I16 and I17 are in. The engine: `packages/studio/src/layers.js`
(the stack, groups, order, `imagePatches`, `anchorOffset`, `reframe`,
`canvasSize`, `imageSize`, `trimBounds`, the corner and rotation
arithmetic), `rimg.js`, and in `index.js` a clear canvas
(`options.background: 'transparent'`) and opacity, blend and lock taken by
every layer type. The window: `renderer/apps/image/` (`layers-panel.js`,
`stage.js`, `insert.js`, `canvas-size.js`, `layered.js`, `words.js`) and
`renderer/apps/edit-in-image.js` for Documents, Worksheets and
Presentations, whose picture goes to Image as a scratch file and comes back
through `win.handBack` and the `picture:returned` event to the window that
opened it only, put in place by a `replacePicture` op of each engine
(width kept, height from the new picture's shape, one undo step). Proven by
`tests/studio-layers.test.js`, `tests/studio-rimg.test.js`,
`tests/edit-in-image.test.js` and the window checks under the key
`image-layers` (`main/verify-image-layers.js`).

What part 1 leaves:

- A corner drag scales a layer uniformly; free transform and perspective are I14 (part 2).
- Text is edited in the Layers panel (a double-click on the canvas goes there), not on the canvas.
- Masks (I10), selections (I11), brushes (I12), clone and heal (I13) are part 2.
- A group has visibility, opacity and a lock; it has no blend of its own (it passes through).
- Draw and Red eye work on the Background picture where it is on the canvas, unturned.
- Edit in Image finds a Worksheets picture on the sheet that is active, and a
  Presentations picture on the slide it was on; a picture whose sheet or slide
  went away is not put back.

#### Part 2 as built (2026-10-12)

I10 to I14 are in. The tools stand in a column down the left of the picture,
as photo editors place them (Move; Rectangle and Ellipse Select, Lasso, Magic
Wand, Select by Colour; Brush, Eraser, Fill, Gradient; Clone Stamp, Healing
Brush; Free Transform, Perspective), with the two colours at the column's
foot, and the chosen tool's options and actions in a bar above the picture.
Picking any tool but Move turns the picture into layers. Right to left, the
column stands on the right.

- **Selections (I11).** A selection is an alpha plane the canvas's size
  (`packages/imaging/src/select.js`): rectangles, ellipses and lassos
  rasterised with smooth edges, the magic wand (tolerance, contiguous) and
  select by colour; New, Add (Shift), Subtract (Alt) and Intersect; Select All,
  Deselect, Invert; Feather, Grow and Shrink by a number of pixels, under the
  options bar's Modify menu so the bar stays one row (an exact distance
  transform, so corners round); marching ants drawn from an outline
  worked out in the image worker. The selection is part of every undo step.
  With a selection, the Adjust and Filters tabs (and the Adjust panel, levels
  and curves included) change only what is selected: the first change makes
  an "Adjusted selection" layer above the chosen layer from the picture under
  it, and every change after makes it again from that same picture, so a
  slider can go back; Deselect returns the tabs to the picture's own quick
  edits. An adjustment layer inserted with a selection takes it as its mask.
  The bucket, the gradient, the brushes and Fill paint only inside it, Delete
  clears it from the chosen layer, Copy to Layer (Ctrl+J) makes it a layer of
  its own in place, and Copy puts it on the clipboard, pasted back where it was.
- **Masks (I10).** `paintMask` on a layer's patch: a bitmap the canvas's size
  kept under `bitmaps/` in the `.rimg`, its place in canvas fractions (moved
  with the layer and by `reframe`), what lies past it (shown or hidden), a
  feather, an inversion and an off switch. The renderer draws a masked layer
  on a surface of its own, cuts it there with the mask (`destination-in`,
  blurred for the feather with the canvas's own edges held out, so only the
  mask's edges soften) and lays it on the frame with the layer's opacity and
  blend; an adjustment layer's surface starts as a copy of the frame, which is
  what it adjusts. Add Mask (Reveal All, Hide All, From Selection), Feather,
  Invert, Turn off, Apply (into a picture's or paint layer's own pixels) and
  Delete are in the Layers panel, the mask's thumbnail beside the layer's;
  a click on it paints the mask (black hides, white shows).
- **Painting (I12, I13).** A stroke is dabs (`packages/imaging/src/paint.js`:
  the brush's fall-off by hardness, a sprite sampled four by four, dabs an
  eighth of the brush apart) laid on a coverage surface, cut to the selection,
  coloured (or filled from the picture moved by the clone's offset) and laid
  on the layer's bitmap through the map from the canvas to it (a turned or
  cornered layer included), live, frame by frame; a pen's pressure narrows
  the dab. Paint goes on the chosen picture or paint layer, on its mask when
  that is chosen, and otherwise on a new paint layer (`kind: 'paint'`, an
  `image` layer over the whole canvas) above it; the Eraser on a layer whose
  own pixels cannot be painted (the Background, which keeps its quick edits,
  or text and shapes) hides through a mask instead. The Fill and the
  Gradient (linear, radial, reflected; to clear or to the second colour) are
  worked in the image worker over the picture as it shows. The Clone Stamp
  copies from where it was Alt+clicked, aligned or from the start each
  stroke; the Healing Brush lays the same copy and then, in the worker, swaps
  its low frequencies for those round the place healed (a weighted blur, so
  the blemish's own colour is left out). A painted region's pixels before and
  after go with the undo step that made them.
- **Free transform and perspective (I14).** `corners` on a picture's patch:
  its four corners in canvas fractions, painted by a projective map
  (`packages/studio/src/warp.js`, bilinear in premultiplied alpha, drawn at
  about the size it lands and kept until the corners change), so a transform
  is never resampled into the project and Reset takes it off. Free Transform's
  corner handles scale from the opposite corner (Shift keeps the shape), skew
  along an edge, distort one corner or pull a perspective; dragging inside
  moves the picture. Perspective sets four pins on something that should be a
  rectangle; Straighten maps them to the rectangle of their average sides,
  the picture with them, and Straighten and Crop cuts the canvas to it; either
  puts the tool down, so no pins stand over the result. Any
  picture layer can be transformed, the Background too.

Proven by `tests/imaging-select.test.js`, `tests/imaging-paint.test.js`,
`tests/studio-warp.test.js`, `tests/studio-masks.test.js` and the window
checks under the key `image-tools` (`main/verify-image-tools.js`).

What part 2 leaves:

- The Background's own pixels are not painted: brushes put paint on a layer
  above it and the eraser hides through its mask, so its quick edits stay
  editable.
- A mask moves with its layer when the layer is dragged, but not when it is
  resized, turned or transformed by its corners.
- Painting on a picture whose corners have been distorted maps each dab
  through the flat map nearest it, so a dab is not itself put in perspective.
- An "Adjusted selection" layer can be changed from the Adjust and Filters
  tabs only while its selection stays and the window is open; saved and
  opened again, it is a picture layer like any other.
- A gradient on a mask sets how much shows from the colours' lightness; there
  is no gradient editor with more than two stops.
- Selections are not saved in the `.rimg`.

### 7.3 Beyond (Image Tier 3, each a declared download at most once)

I18 background removal with a small segmentation model; I19 content-aware
fill by patch matching (no model); I20 RAW (CR2, NEF, ARW, DNG) with
exposure and white balance; I21 HDR merge and panorama; I22 a recorded
action replayed over a folder.

## 8. The screen recorder

Recorded on this computer, never uploaded. Lives in Video (Record) and is
offered from Presentations (Insert → Screen Recording) and the launcher.

| # | Item | Notes |
|---|---|---|
| R1 | Sources: a whole screen, one window, or a region drawn on a screen | Region recorded as the full screen and cropped by the layer, so the crop can change after; `shell.capture.sources()` |
| R2 | The camera, at the same time, as a second file (D6) | `cameo.js`'s patient open moved to a shared module; choose the camera, its resolution and a mirror |
| R3 | The microphone, with a level meter before and during | Choose the device; a noise gate on the take is an export option, never destructive |
| R4 | The computer's own sound (loopback) | Main process: `session.setDisplayMediaRequestHandler` with `audio: 'loopback'` on Windows, a contract entry in `office-shell`; macOS needs a different route, checked when that platform ships |
| R5 | A live preview with the camera inset where it will sit; the inset dragged to a corner before recording | |
| R6 | Countdown (3 s, a setting), pause and resume, stop from a hotkey (a global shortcut while recording) and from the tray | `globalShortcut` in the main process, released on stop |
| R7 | Mouse: the pointer kept or hidden, clicks shown as rings, and a highlight round the pointer | Pointer position recorded as data beside the video (main process polls `screen.getCursorScreenPoint` at the frame rate), drawn by a layer, so it can be changed after |
| R8 | Keystrokes shown on screen | Optional; recorded as data, drawn by a layer, off by default and never for password fields |
| R9 | Draw on the screen while recording (pen, arrow, highlighter) | A transparent always-on-top window; strokes recorded as data and drawn by a layer |
| R10 | Quality settings: frame rate (15, 30, 60), resolution cap, bitrate | Defaults that keep a 1080p screen at 30 fps under about 8 Mbps |
| R11 | Streaming to scratch files a few megabytes at a time, so a long take never sits in memory | `shell.fs.temp`, `append`, `dropTemp`, as `screen-record.js` does |
| R12 | A take survives a crash: the scratch files are offered back on the next start | |
| R13 | Disk space checked before and during; a warning at five minutes left | |
| R14 | On stop, the takes land on the timeline: screen on lane 1, camera above as an inset, sound on its own lane, pointer and keystroke layers above | A recording from Presentations goes onto the slide instead, as today |
| R15 | Retakes: record again over a range of the timeline | |
| R16 | Recording a single window keeps recording it when it moves or is covered | Chromium's window capture does; check minimised windows |

## 9. Introductory and explainer videos

The job: someone sits in front of their screen and camera and makes a short
video introducing themselves, a product or a how-to, then sends it. It
needs the recorder (8) and the editor (6) joined by a guided flow. Entry:
Video → New → Introduction video, and a card in the launcher's Start
something.

| # | Item | Notes |
|---|---|---|
| X1 | A template of scenes: title card, talking head (camera full frame), screen with the camera in a corner, side by side, end card | Each scene a studio document fragment with layout, transitions and text placeholders |
| X2 | Layouts switchable per scene after recording, since screen and camera are separate layers | D6 is what makes this possible |
| X3 | A script per scene, shown as a teleprompter above the camera preview while recording, scrolling at a set speed | Kept in the project; never sent anywhere |
| X4 | Record scene by scene, retake any scene, reorder scenes | Each scene's takes kept until the project is cleaned up |
| X5 | Automatic zoom on the area the pointer works in, from the pointer data (R7) | Keys on the screen layer's zoom and pan, editable after |
| X6 | Brand kit: logo, colours, fonts and a lower third with name and role, applied to every scene | Kept per user in the app's settings |
| X7 | Background music from a small built-in set, ducked under speech | Licence-clean tracks only, shipped with the app |
| X8 | Captions burnt in or kept as a track (V23, V24) | |
| X9 | Silence and filler trimming: gaps over a threshold found from the sound level and offered for removal | Level based first; words from a local speech model later |
| X10 | Background blur or replacement for the camera | Tier 3: the segmentation model of I18 |
| X11 | Export presets by destination: YouTube 16:9 1080p, LinkedIn and X square or 16:9, Reels, Shorts and TikTok 9:16, email-friendly small MP4 | Uses 11 |
| X12 | A thumbnail picture made from a chosen frame with the title on it | Opens in Image |

## 10. Performance

Budgets are tests, as `tests/perf-budget.perf.js` holds the documents to
theirs. Every item below gets a budget once it lands.

| # | Item | Notes |
|---|---|---|
| P1 | Decode with WebCodecs `VideoDecoder` from demuxed samples, rather than seeking `<video>` elements | Frame-exact seeking and scrubbing; a demuxer for MP4 and WebM in `packages/media` (the probe already parses boxes and elements) |
| P2 | A frame cache around the playhead and a thumbnail cache per clip, with a memory cap | Frames as `VideoFrame` or `ImageBitmap`, closed when evicted |
| P3 | Composite on the GPU: WebGL2 for layers, blend modes, colour, keying and LUTs | The 2D canvas path stays as the reference; golden-frame tests compare the two |
| P4 | Paint and encode in a worker with `OffscreenCanvas`, so the window never stutters while exporting | A worker entry in `build/bundle.js` |
| P5 | Preview at reduced resolution while playing, full resolution when paused | A setting: Auto, Full, Half, Quarter |
| P6 | Proxies: a small copy of each 4K or long source made in the background and used for editing; the original used on export | Kept in the project's cache folder |
| P7 | Export faster than real time (WebCodecs encode, D3), in the background, with progress, time left, pause and cancel | The window stays usable; a finished export is announced |
| P8 | Hardware encoding where the platform offers it (`hardwareAcceleration: 'prefer-hardware'`), with software fallback | |
| P9 | Waveforms computed once in a worker and cached | |
| P10 | Large stills decoded at the size they are shown | `createImageBitmap` with resize options |
| P11 | Budgets: a 10-minute 1080p project with 50 clips opens in under 3 s, scrubs at 30 frames a second on a mid-range laptop, and exports at least twice real time with hardware encoding; a 40-megapixel photo opens in Advanced in under 2 s | Written as tests where the engine allows, window checks otherwise |
| P12 | Memory: the video window stays under 1.5 GB on the budget project | As the documents' budget does |

## 11. Formats

What opens and what is written. "Platform" means the codec comes from
Chromium and the operating system in Electron and must be checked on each
platform before it is promised; a spike in Phase 0 lists what Electron 43
on Windows decodes and encodes through WebCodecs.

### 11.1 Opening

| Kind | Formats |
|---|---|
| Video | MP4 and MOV (H.264; H.265 and AV1 where the platform decodes them), WebM (VP8, VP9, AV1), MKV, M4V; AVI and WMV are named as not supported, with a plain sentence |
| Sound | WAV, MP3, M4A and AAC, Ogg Vorbis, Opus, FLAC |
| Pictures | PNG, JPEG, WebP, GIF (animated as a clip), AVIF, BMP, SVG; HEIC and TIFF decoded in the app's own code (Image Tier 1) |
| Captions | SRT, WebVTT |
| Projects | `.rvid`, `.rimg` |

### 11.2 Saving video

| Format | Codecs | Notes |
|---|---|---|
| MP4 | H.264 and AAC | The default; plays everywhere. Our own muxer (D3) |
| MP4 | H.265 (HEVC) and AAC | Where the platform encodes it; smaller at the same quality |
| MP4 or WebM | AV1 and Opus | Where the platform encodes it |
| WebM | VP9 and Opus | The free-codec choice; available everywhere Chromium is |
| MOV | H.264 and AAC | The same samples in a QuickTime box, for editors that ask for MOV |
| GIF and animated WebP | | Short clips; palette per frame for GIF |
| PNG sequence | | Frames in a folder, for other tools |

Settings for every video export: resolution (480p, 720p, 1080p, 1440p, 4K,
or the project's own), frame rate (24, 25, 30, 50, 60, or the source's),
quality as a bitrate or a target size, and the destination presets of X11.
Chapters from markers (V7) in MP4. Captions as a separate SRT or WebVTT
file, or as a soft subtitle track in MP4 and WebM.

### 11.3 Saving sound

WAV (already written by `capture/media-encode.js`), M4A (AAC), Ogg Opus, and
MP3 if an encoder of our own or an AGPL-compatible one is added; loudness
normalisation (V30) as an option.

### 11.4 Saving pictures

PNG, JPEG (quality), WebP (quality, lossless), AVIF where the platform
encodes it, BMP, TIFF (written by our own code), ICO for icons, PDF (one
page, through `@rutba/pdf`), and a flattened PSD; metadata kept or stripped;
a size cap and copy to the clipboard. `.rimg` keeps everything editable.

## 12. Phases and releases

Each phase is releasable on its own and follows docs/RELEASING.md (a
feature takes the minor number, the note names its checks).

| Phase | Contents | Done when |
|---|---|---|
| 0. Spike | What Electron 43 decodes and encodes through WebCodecs on Windows (11); the main-process loopback handler (R4); the timeline model decision (D1) written down | A table in this file, D1 to D7 taken |
| 1. Video on the engine | V1 to V10, V13, V15, V21, V22, V26, V32, V33; export with sound through `renderVideo` | Window checks open two clips and a still, split, trim, add a title and a crossfade, export, and read the file's duration, a frame's pixels and its audio track; engine tests for split and ripple; gate green |
| 2. The recorder | R1 to R6, R10 to R14; Presentations' screen recording moved onto it | Window checks with `RUTBA_FAKE_MEDIA=1` and a fake screen record both streams, land them as two layers and export |
| 3. Introductory videos | X1 to X4, X6, X8 (import), X11, R7, X5 | A check makes a three-scene intro from fake sources and exports it at two presets |
| 4. Faster and more formats | P1, P2, P4, P5, P7, P8, P9; D3's MP4 muxer; 11.2 MP4, WebM, MOV, GIF; 11.3 | Budgets P11 and P12 as tests; exported files read back by `probeMedia` and a decoder |
| 5. Image Quick and Advanced | I1 to I17 | Golden-image engine tests per operation; window checks for the layers panel, a mask, a selection and Edit in Image |
| 6. Tracks and colour | V11, V12, V14, V16 to V19, V23, V27 to V29, V34 to V36 | Checks for picture in picture, keyframes and captions in and out |
| 7. Beyond | V20, V24, V30, X7, X9, X10, X12, R8, R9, R15, R16, I18 to I22, P3, P6, 11.4 remainder | Each a declared download at most once, each with its own check |
| 8. Ready for a web host | No window in Video or Image touches anything outside `shell` and the engine; the engine tests run in a browser page; a list here of what a web `shell` must provide | |

The recorder comes before tracks because the introductory video depends on
it and Presentations already has its pieces; the owner may reorder (14).

### Phase 1 as built (2026-10-12)

The Video app is `apps/desktop/renderer/apps/video/` (index, timeline,
stage, inspector, ribbon, media, export, css) on `@rutba/studio`: the
project model in `src/project.js`, the file in `src/rvid.js`, the engine's
`driveVideoElements` shared by the preview and the export. V1 to V10, V13,
V15, V21, V22, V26, V32 and V33 are in, with export with sound through
`renderVideo`. Proven by `tests/studio-project.test.js`,
`tests/studio-rvid.test.js` and the window checks under the key
`video-editor` (`main/verify-video-editor.js`), which make two clips with a
tone each and a still at run time and read the exported file back with
`probeMedia` (whose WebM and fragmented MP4 lengths and track lists came
with this phase).

What Phase 1 leaves, for the phases that own it:

- Media is read whole into a Blob per source (D4 as written); a two-hour
  recording sits in memory. A range-reading transport belongs with P1 and P2.
- Thumbnails are one a second, at most 120 per source, made by seeking a
  `<video>`; a still's thumbnail does not show its turn or crop.
- J plays backward by seeking, without sound; speed and reverse as edits are
  V11.
- Volume runs from silent to the clip's own level (the mixer clamps at 1); a
  boost needs the engine's gain stage to allow more.
- A track's volume is set from the properties of a clip on it, not on the
  track's header.
- Titles are `text` layers (and a shape behind a lower third or an end card);
  word art and the rest of the insert catalogue on the timeline are later.
- Export is real time through `MediaRecorder` (MP4 where the machine records
  it, WebM otherwise); the encoder of our own is D3, Phase 4.

### Phase 2 as built (2026-10-12)

The recorder is `apps/desktop/renderer/apps/video/recorder/`: `panel.js`
(the Record box), `session.js` (the sources, the sound graph and the
recorders), `camera.js` (the shared, patient camera open, moved out of
Presentations' `cameo.js`) and `css.js`. Its arithmetic is
`packages/studio/src/recorder.js`: the settings and what they cost the
disk, the camera's inset and the corner it settles in, a drawn region as a
crop, the pausing clock, the writer that streams a recording to its file
(four megabytes a message at most), and `placeTakes` and `settleTakes`,
which put the takes on the timeline and settle a recovered recording to
its files' real lengths. R1 to R6 and R10 to R14 are in:

- R1: Screen, Window or Region, from `shell.capture.sources()` (screens now
  carry their size). A region is drawn on the live picture and kept as the
  screen layer's crop, changeable after in the clip's properties. A region
  can be drawn on a window as well as a screen.
- R2: the camera at 480, 720 or 1080 lines, which camera, mirrored or not,
  cut to a circle, a rounded or a square rectangle, small, medium or large.
- R3, R4: the microphone and the computer's sound each with a level meter
  before and during. The computer's sound comes through
  `capture.choose`: the main process answers the window's
  `getDisplayMedia` with the source it chose and, on Windows, `audio:
  'loopback'` (`session.setDisplayMediaRequestHandler`); a request no
  window chose for is refused.
- R5: the frame as it will be recorded, with the camera's inset where it
  will sit, dragged to a corner.
- R6: a countdown (none, 3, 5 or 10 seconds; 3 by default), Pause and
  Resume, and `capture.controls`: Ctrl+Shift+F10 stops and Ctrl+Shift+F9
  pauses from anywhere while recording, a tray icon offers both, and the
  window can minimise itself while recording; all of it is let go on stop
  or when the window closes.
- R10: 15, 30 or 60 frames a second, a size held to 720, 1080 or 1440 on
  the shorter side (or the source's own), and three qualities as bits per
  pixel; the default keeps a 1080p screen at 30 fps near 6 Mbps.
- R11: the takes go to `Videos/Rutba Recordings` (a setting) through
  `capture.take` and `fs.append`; Presentations' through `fs.temp`.
- R12: every five seconds while recording, the project with its takes so
  far (marked unfinished) is the window's recovery copy; the launcher's
  Recover opens it in Video, settled to the files' real lengths.
  Presentations' recording is copied the same way.
- R13: `capture.space` before and every five seconds during; a warning
  under five minutes left, and a stop that keeps the take under twenty
  seconds.
- R14: the screen on the first free picture lane, the camera on the lane
  above, each sound on a free sound lane, after what is already there.
  Presentations' Insert, Screen Recording opens the same box without the
  camera and puts the screen's take on the slide.

Proven by `tests/studio-recorder.test.js` and the window checks under the
key `recorder` (`main/verify-recorder.js`), which record one of the
suite's own windows with Chromium's stand-in camera and microphone, land
the takes, export, read the files back with `probeMedia`, and recover a
copy made while recording; `screenrec` holds Presentations to the same
recorder. A check run lists only the suite's own windows, and is handed
the global key's press rather than taking a key from the keyboard.

What Phase 2 leaves:

- The computer's sound is proven only as far as the main process answering
  `getDisplayMedia`; a check never records the machine's own output. On
  macOS there is no loopback: the box says so.
- The global keys and the tray are not exercised for real by the checks
  (no key is taken from the person's keyboard and no icon put on their
  desktop); the press is handed in.
- A region is drawn on the live picture with the camera's inset over it,
  and the cropped result is seen on the timeline, not in the box.
- Moving or covering a recorded window is Chromium's (R16); a minimised
  window sends no frames.
- The pointer, keystrokes and drawing on the screen (R7 to R9) and
  retakes (R15) are later phases.
- Takes are WebM (VP9 or VP8 and Opus), so a take cut short by a crash
  still plays; Presentations' recordings stay MP4 where the machine makes
  it, as before.

### Phase 3 as built (2026-10-12)

Introductions are a video project with an `intro`: its scenes in order, its
brand kit, the template's words in the language it was made in, and the
teleprompter's speed. The arithmetic is `packages/studio/src/intro.js`
(scenes, the template, layouts, brand kit, recording and retakes, the
teleprompter), `captions.js` (SubRip and WebVTT), `pointer.js` (the pointer
layer and the automatic zoom) and `presets.js` (export by destination). The
window is `apps/desktop/renderer/apps/video/intro/` (the Scenes panel, the
brand kit, their styles), with the teleprompter in the recorder's
`panel.js`. Entry: the launcher's Introduction video card, New Introduction
Video in Video's menu and on its Home ribbon.

- X1: a new introduction is a title card, a talking head, the screen with
  the camera in a corner and an end card; side by side is the fifth layout.
  A scene is a fragment of the timeline: its layers carry `scene` and
  `role` (ground, screen, pointer, camera, logo, band, title, subtitle, the
  take's sounds), each role on a lane of its own, with fades between scenes.
  `composeIntro` makes every scene's layers again from the scenes after
  reading back what the timeline says about each (its card's words, its
  lower third's words, how its take was trimmed, its zooms, its pointer's
  look, the camera's corner); everything that is not a scene's (captions,
  anything added) is left where it is.
- X2: the layout of a recorded scene is switched from its card or the
  Introduction tab and its take is laid out again: a talking head keeps the
  screen's take hidden, so switching to the screen with the camera brings it
  back. Each layout is written for 16:9, square and 9:16, and an export to
  another shape lays every scene out for it (screen above the camera in
  9:16), leaving the project as it was.
- X3: each recorded scene has a script, shown in the recorder above the
  frame as a teleprompter that scrolls at 110, 140 or 170 words a minute
  (Slow, Normal, Fast, kept in the project) from a second after recording
  starts, the line being read a third of the way down.
- X4: Record (or Retake) in a scene's card opens the recorder with the
  camera on; the take lands in its scene and the scenes after it move on.
  Every take is kept with its scene and any can be put back with Use;
  scenes move earlier or later, are added after the one picked, and are
  taken out (their files stay on the disk).
- X6: the brand kit (a name and role for the lower third, a logo, three
  colours, a font, the lower third and a corner logo each on or off) is kept
  per person under `video.brand` and drawn into every scene: the ground in
  its background colour, the cards' words in its colours and font with the
  logo above them, a title card still in the template's words taking the
  brand's name and role, the lower third over the first recorded scene with
  a camera (right to left for an Urdu or Arabic name, and kept out of the
  camera's way: on the side the name starts from over a talking head,
  across from an inset in a bottom corner, over the camera's box side by
  side, and between the screen and the camera in 9:16), and the logo in a
  corner of recorded scenes away from the camera. The box shows a title
  card and a lower third as they will look.
- X8 (import): Insert, Import Captions reads SubRip and WebVTT (the header,
  notes, styles, cue settings, voices, tags and entities handled; a cue that
  overlaps the one before goes a lane up) onto a track above every other
  picture, each cue a `text` layer marked `caption`, drawn on the scrim's
  pill, in two lines where it has two, right to left in Urdu and Arabic.
  Export Captions writes the track as .srt or .vtt. A caption is a clip, so
  it moves, trims, splits and has its words edited like a title. The
  engine's text layers gained several lines and a font family of their own.
- X11: Export's Made for offers the project as it is, YouTube (1920 by 1080,
  30 fps, 10 Mbps), LinkedIn and X at 16:9 or square, Reels, Shorts and
  TikTok (1080 by 1920) and email (480 lines in the project's shape, 24 fps,
  about a megabyte for eight seconds), with the frame, the likely size, and
  a word when it is too long or too big for where it goes. The choice is
  remembered. A plan takes a frame size and bitrates of its own for this.
- R7: while Video records, the main process writes the pointer down
  (`capture.pointer`): polled with `screen.getCursorScreenPoint` at the
  frame rate over a screen (mapped by its display's bounds) or one of the
  suite's own windows (by its bounds), the paused time left out, and a press
  in the suite's own windows. It lands as a `pointer` layer on the lane
  above the screen, in the recording's group, drawing a highlight round the
  pointer and a ring where it clicked wherever the screen layer draws that
  point (through its crop, its fit, an inset's box and its zoom at that
  instant); its look (highlight, rings, a larger drawn pointer, size,
  colour) is changed after in its properties. A split or a trimmed start
  keeps its keys and its zooms in the clip's own time.
- X5: Zoom on Clicks (Clip tab, and the screen's properties) finds runs of
  clicks close together and zooms the screen on each, following the clicks,
  as zoom and pan keys made from a list kept on the layer (`autoZoom`);
  each zoom is made closer or wider or taken away there. An introduction's
  screen scene is zoomed this way as soon as it is recorded.

Proven by `tests/studio-captions.test.js`, `studio-presets.test.js`,
`studio-pointer.test.js` and `studio-intro.test.js`, and the window checks
under the key `intro-video` (`main/verify-intro-video.js`): captions in four
languages imported, drawn, edited and exported; the presets' frames and an
email export read back; a suite window recorded with the pointer handed its
moves and clicks, its highlight and ring on the stage, a zoom made, changed
and taken away; and an introduction from the launcher's card with a brand
kit, the end card taken out, a script on the teleprompter, the talking head
and the screen scene recorded with the stand-in camera, the talking head
taken again, the scenes reordered and the screen scene switched to side by
side, a caption file imported, and exports for YouTube and Reels read back
with `probeMedia` for their size, length and tracks.

What Phase 3 leaves:

- A press is seen only in the suite's own windows: Electron offers no
  global mouse hook without a native module, so a click in another program
  is not recorded (its pointer is, over a screen). Over another program's
  window the pointer cannot be mapped, since only the suite's own windows
  report their place; the recorder says so. The checks hand the pointer its
  positions and presses; they never read the person's mouse.
- The pointer's own arrow in the recording cannot be hidden (Chromium's
  screen capture draws it); Larger pointer draws one of ours over it.
- The automatic zoom zooms on runs of clicks, not on a pointer that only
  moves or rests, nor on typing.
- A layout switch, a retake, a move or a brand change makes the scenes'
  layers again: what is read back is listed above; any other change made on
  the timeline to a scene's own layers (a title dragged, a colour changed on
  one layer) goes.
- A scene's recording is not copied for recovery while it records, as a
  plain recording is (R12).
- Captions are imported as a track and burnt in on export; exported as a
  soft subtitle track inside the file (11.2) and from speech (V24) are later.
- The recorded takes of a scene are kept until the project is cleaned up,
  but there is no Clean Up command yet (`takeFiles` lists them).
- A take put back with Use is zoomed afresh on its own clicks; changes made
  to the zooms of the take it replaces are not kept with that take.
- The scene card's script box keeps its words when it loses the focus;
  Record keeps them too, but closing the window mid-word does not.

## 13. Conventions that bite

- **A new workspace package** is declared in `apps/desktop/package.json`
  or the installer misses it (`tests/desktop-deps.test.js`); `@rutba/studio`
  already is.
- **Tests** go in the root `tests/` as `*.test.js`; tests inside a package
  are not found.
- **Words**: every visible word through `t()`, with Urdu in
  `catalogues/ur.js`; `npm run` the message extract; `tests/messages.test.js`,
  `catalogues.test.js` and `english-left.test.js` must pass. An engine
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

## 14. Open questions for the owner

- Should the recorder's default be the two-file recording (D6) or one
  composited file?
- Is Advanced a mode inside Image, or a separate app on the launcher?
  (Taken as a mode inside Image, 7.2; the owner may still decide otherwise.)
- Do Image and Video keep their names when they become editors? (Image keeps
  its name for now.)
- Section 12 puts the recorder and introductory videos (Phases 2 and 3)
  before tracks and colour (Phase 6). Is that the order wanted?
- Which destination presets matter most for introductory videos (X11)?
- Should MP3 export (11.3) wait for an encoder of our own, or is an
  AGPL-compatible library acceptable?
- Background music (X7): which licence-clean tracks may ship with the app?
