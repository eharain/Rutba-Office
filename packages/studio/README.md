# @rutba/studio

The renderer the Video app's editor, the Image app's advanced editor and the
recorder are built on. It came from the consumer suite's Studio
(`consumer/studio/packages/editor-core` and the pure parts of
`consumer/studio/apps/studio/lib`, at consumer `1aebcb67`, 2026-10-09). The
plan for building on it is [docs/STUDIO.md](../../docs/STUDIO.md).

## What it is

One pipeline, from a document to pixels:

```
doc (patches)  ->  buildPlan  ->  applyLayerPatches  ->  compileLayers  ->  paintFrame(ctx, plan, t)
                                                                          ->  renderVideo  (MediaRecorder)
                                                                          ->  renderImage  (one frame)
```

- **A document** (`src/doc.js`) holds the recipe and nothing derived: the
  pictures it references and a list of layer patches. Geometry is fractional
  of the frame, so the size is a render decision. An image is a one-frame
  video, so both editors share the format.
- **A layer** has a type, a z order, `timing` (start, end), enter and exit
  ramps (fade, slide, push, zoom, wipe), keyframes on position, size,
  opacity, rotation, zoom and pan (linear or eased), a mask (rectangle,
  ellipse or polygon, invertible), a blend mode, crop, zoom and pan.
- **Layer types**: photo, video (with a source offset), gradient, caption,
  title, image, text, qr, outro, progress, edges, sound (volume, fades,
  loop, mix, duck or solo); and, registered by the toolset, word art (with
  warps), shapes (30 geometries), adjustment layers (brightness, contrast,
  exposure, levels, curves, temperature, tint, saturation, vibrance,
  vignette, sharpen, posterize, invert, with presets), icons, lists and
  charts.
- **Editing primitives**: `layerBounds`, `hitTestLayers`, `layerHandles`,
  `hitTestHandles`, `scaleFromDrag`, `resizePatch`, `withLayerStateAt`.
- **Output**: `renderVideo` records the canvas and the mixed sound in real
  time through `captureStream` and `MediaRecorder` (MP4 H.264 where the
  platform has it, WebM otherwise); `renderImage` and `renderImageSet` write
  PNG, JPEG or WebP.

## How a window uses it

Import from `@rutba/studio/host`, never from the engine files one by one.
Importing `host` registers the toolset; `paintFrame` skips a layer type it
has no painter for rather than throwing, so a window that forgets it shows
word art in the preview and loses it in the export.

The engine has no media transport. A window gives it one, once:

```js
import { configureMediaFetch } from '@rutba/studio/host';
configureMediaFetch(async (url, { signal, range } = {}) => /* a Blob of the bytes */);
```

The answer must be a Blob, so the canvas that draws it is never tainted and
`captureStream` and `getImageData` keep working.

## Layout

| Path | What |
|---|---|
| `src/index.js` | Plan, compile, paint, media loading, hit testing, sound, `renderVideo` |
| `src/image.js` | Stills: `buildImagePlan`, `renderImage`, `renderImageSet`, `IMAGE_SIZES` |
| `src/toolset.js` | Registers word art, shapes, adjustments, icons, lists, charts |
| `src/wordart.js`, `shapes.js`, `adjust.js`, `icons.js`, `textblock.js`, `charts.js` | The toolset's layer types |
| `src/resolve.js` | What a stored document must fetch; audio clips by injection |
| `src/lib/` | `aspects.js`, `beats.js` (snapping to a music track's beats), `storyboard.js` |
| `src/host.js` | The one import a window uses |
| `src/doc.js` | The project document: `emptyDoc`, `normaliseDoc`, `addLayer`, `updateLayer`, `setOption` |
| `src/zoom.js` | Stage zoom and pan arithmetic |
| `src/insert-catalog.js` | Everything that can be put on a creative, arranged for a panel |
| `src/chart-data.js` | A chart layer's data as text and back |
| `src/capture/media-encode.js`, `audio-edit.js` | WAV encoding, trimming and mixing of recorded sound |
| `port/` | The Studio editor's React UI, to be rewritten onto `@rutba/office-ui` (see its README) |

## Tests

`tests/studio-*.test.js` in the repository root, run by `npm test` (173
tests on landing): painting, toolset, icons, lists, charts, resolve,
recording, settings, insert and zoom. Two of them read `port/` files as
text; when a component is ported they move with it.
