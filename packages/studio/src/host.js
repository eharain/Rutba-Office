/**
 * The studio renderer as a window reaches it, with the toolset registered.
 *
 * Office's counterpart of the consumer app's `lib/renderer.js` (kept under
 * `port/renderer.js` for reference). Two things differ from that file:
 *
 * 1. No media transport is installed here. The engine ships with none on
 *    purpose: a page reads bytes the way its host allows. A window passes its
 *    own to `configureMediaFetch` once, at start (in the desktop app, the
 *    `rutba://file/` protocol through `shell.fs`), and the canvas stays clean
 *    because the bytes come back as a Blob.
 *
 * 2. Importing this module REGISTERS word art, shapes, adjustment layers,
 *    icons, lists and charts into the renderer. paintFrame skips a layer type
 *    it has no painter for rather than throwing, so a window that imports the
 *    engine without this module gets a headline that previews and is missing
 *    from the export. Every window reaches the renderer through here.
 *
 * The surface is named rather than star-exported: `toolset.js` re-exports two
 * names from `index.js`, and an ambiguous star export is silently dropped.
 */
import { loadImages as loadImageSet } from './index.js';

export {
  // shapes of the output
  ASPECTS, THEMES, DEFAULTS,
  // compile and paint
  buildPlan, applyLayerPatches, compileLayers, paintFrame,
  // media
  loadImage, releaseImages, loadVideo, releaseVideos,
  loadAudioTrack, audioContext, isAudioFile, isVideoFile,
  configureMediaFetch,
  // editor primitives
  layerBounds, hitTestLayers, layerHandles, hitTestHandles, scaleFromDrag, resizePatch,
  withLayerStateAt,
  // video output
  renderVideo, startAudioPreview, soundLayers, clipFromSoundLayer,
  pickMimeType, extensionFor, unsupportedReason, videoFileName,
  // the extension registry
  registerLayerType, registeredLayerTypes,
} from './index.js';

/**
 * Load pictures and answer with the entries alone, an array, which is what
 * `buildPlan({ images })` indexes. The engine's own `loadImages` answers
 * `{ images, failures }`; passing that report on draws no picture at all and
 * says nothing, so this wrapper is the one way windows load them.
 */
export async function loadImages(urls, opts) {
  const { images } = await loadImageSet(urls, opts);
  return images;
}

export { buildImagePlan, renderImage, renderImageSet, supportedImageFormats, IMAGE_SIZES }
  from './image.js';

export { mediaManifest, resolveAudioClips, trackUrl } from './resolve.js';

export {
  WORDART, SHAPE, ADJUST,
  WORDART_PRESETS, WARPS, GEOMETRIES, ADJUST_PRESETS,
  ADD_MENU, applyAdjustment, wordArtResize,
  SHAPE_GROUPS, GEOMETRY_LABELS, paintShapeTile,
  ICONS, ICON_NAMES, ICON_GROUPS, ICON_LABELS,
  LIST_KINDS, LIST_LABELS,
  CHART_KINDS, CHART_LABELS, CHART_PALETTE,
} from './toolset.js';
