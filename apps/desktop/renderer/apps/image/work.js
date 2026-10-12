// The window's side of the image worker (worker.js): one worker for the
// window's life, jobs answered by id. Where a worker cannot be started (a
// web host that does not serve it), the pixel work runs here instead, so the
// picture is still right, only slower.
import { applyPixels } from '@rutba/imaging/pixels';
import { runTool } from './tool-jobs.js';

let worker = null;
let failed = false;
let next = 1;
const waiting = new Map();

function start() {
  if (worker || failed) return worker;
  try {
    worker = new Worker('image-worker.js');
    worker.onmessage = (e) => {
      const job = waiting.get(e.data?.id);
      if (!job) return;
      waiting.delete(e.data.id);
      if (e.data.error) job.reject(new Error(e.data.error));
      else job.resolve(e.data);
    };
    worker.onerror = (e) => {
      e.preventDefault?.();
      // A worker that will not load fails every job now waiting; the next
      // job runs in the window.
      failed = true;
      worker = null;
      for (const job of waiting.values()) job.reject(new Error(e.message || 'the image worker stopped'));
      waiting.clear();
    };
  } catch {
    failed = true;
    worker = null;
  }
  return worker;
}

function post(message, transfer = []) {
  const w = start();
  if (!w) return null;
  const id = next++;
  return new Promise((resolve, reject) => {
    waiting.set(id, { resolve, reject });
    w.postMessage({ ...message, id }, transfer);
  });
}

/**
 * Tone and detail applied to an ImageData, off the window's thread. Answers
 * a new ImageData; the one handed in is given up to the worker.
 */
export async function pixelsOff(imageData, settings, spots = []) {
  const { width, height } = imageData;
  const copy = new Uint8ClampedArray(imageData.data);
  const asked = post({ kind: 'pixels', width, height, buffer: copy.buffer, settings, spots }, [copy.buffer]);
  if (asked) {
    try {
      const { buffer } = await asked;
      return new ImageData(new Uint8ClampedArray(buffer), width, height);
    } catch {
      /* the window does it */
    }
  }
  const img = { data: new Uint8ClampedArray(imageData.data), width, height };
  applyPixels(img, settings, { spots });
  return new ImageData(img.data, width, height);
}

/**
 * One picture of a batch, decoded, resized, encoded and its metadata kept or
 * dropped, in the worker: `{ bytes, width, height, format, quality, fits }`.
 */
export async function convertOff(bytes, options) {
  const copy = new Uint8Array(bytes);
  const asked = post({ kind: 'convert', bytes: copy, options }, [copy.buffer]);
  if (!asked) throw new Error('the image worker could not start');
  return asked;
}

/**
 * A pixel job of the advanced editor (tool-jobs.js) off the window's thread:
 * selections, fills, gradients, healing. `job` carries ArrayBuffers, which
 * go to the worker as copies (the window keeps its own); answers what the
 * job answers. Run in the window where the worker cannot start.
 */
export async function toolOff(job) {
  const copy = {};
  const transfer = [];
  for (const [k, v] of Object.entries(job)) {
    if (v instanceof ArrayBuffer) {
      const c = v.slice(0);
      copy[k] = c;
      transfer.push(c);
    } else copy[k] = v;
  }
  const asked = post({ ...copy, kind: 'tool' }, transfer);
  if (asked) {
    try {
      return await asked;
    } catch {
      /* the window does it */
    }
  }
  return runTool(job).result;
}
