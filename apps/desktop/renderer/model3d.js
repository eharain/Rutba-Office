// Insert → 3D Models, in the window: a model file read with whatever lies
// beside it, packed as one .glb, its pictures decoded by the page, and the
// model drawn into a PNG by the suite's own renderer (@rutba/imaging/model3d)
// at the view asked for — what the document keeps as the model's picture.

import { readModel, toGlb, gltfFiles, renderModel, fit } from '@rutba/imaging/model3d';

/** The longest side of a model's picture when it is first put in, in pixels. */
export const MODEL_PICTURE = 900;

/** The view a model is put in at, and Reset 3D Model goes back to: turned a little, seen a little from above. */
export const DEFAULT_MODEL_VIEW = { yaw: 25, pitch: 15, roll: 0 };

const folderOf = (p) => String(p).replace(/[\\/][^\\/]*$/, '');

/**
 * A model file as one self-contained .glb. A .gltf's buffers and pictures
 * are read from beside it — a plain name in its own folder or below, never
 * a path that climbs out of it or starts somewhere else, so a model cannot
 * carry another file of the computer's into a document.
 */
export async function loadModelFile(shell, file) {
  const { bytes } = await shell.fs.read({ path: file });
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const files = new Map();
  for (const uri of gltfFiles(u8)) {
    if (/^[a-z]+:/i.test(uri) || /^[\\/]/.test(uri) || uri.split(/[\\/]/).includes('..')) throw new Error(`The model names a file outside its folder ("${uri}"), which is not read`);
    const got = await shell.fs.read({ path: `${folderOf(file)}/${uri}` }).catch(() => null);
    if (got) files.set(uri, got.bytes instanceof Uint8Array ? got.bytes : new Uint8Array(got.bytes));
  }
  return toGlb(u8, { readFile: (u) => files.get(u) || null });
}

/** A model's pictures, decoded by the page and no larger than `maxSide`: `Map(index → { width, height, data })`. */
export async function decodeTextures(model, maxSide = 1024) {
  const out = new Map();
  await Promise.all((model.images || []).map(async (img, i) => {
    if (!img?.bytes?.length) return;
    try {
      const bitmap = await createImageBitmap(new Blob([img.bytes], { type: img.mimeType || 'image/png' }));
      const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
      const width = Math.max(1, Math.round(bitmap.width * scale));
      const height = Math.max(1, Math.round(bitmap.height * scale));
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      ctx.drawImage(bitmap, 0, 0, width, height);
      out.set(i, { width, height, data: ctx.getImageData(0, 0, width, height).data });
      bitmap.close?.();
    } catch {
      // A picture the page cannot read is left out; the base colour is drawn alone.
    }
  }));
  return out;
}

/** Pixels as PNG bytes. */
export async function pngOf({ width, height, data }) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  canvas.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(data.buffer, data.byteOffset, data.length), width, height), 0, 0);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

/** A data: address for pixels, for a quick look while a model is turned. */
export function urlOf({ width, height, data }) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  canvas.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(data.buffer, data.byteOffset, data.length), width, height), 0, 0);
  return canvas.toDataURL('image/png');
}

/**
 * A model ready to draw again and again: the .glb read and its pictures
 * decoded once. `draw({ view, width, height })` gives the pixels;
 * `size(view, longest)` the picture's size for the model at that view.
 */
export async function modelDrawer(glb) {
  const model = readModel(glb);
  const textures = await decodeTextures(model);
  return {
    model,
    triangles: model.triangles,
    size: (view, longest = MODEL_PICTURE) => fit(model, view, longest),
    draw: ({ view, width, height, samples = 2 }) => renderModel(model, { width, height, view, textures, samples }),
  };
}
