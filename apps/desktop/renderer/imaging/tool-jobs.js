// The advanced editor's pixel jobs (docs/STUDIO.md 7.2, I10 to I13): the
// magic wand and select by colour, a selection feathered, grown or shrunk,
// its outline for the marching ants, the bucket fill, the gradient and the
// healing brush's blend. Run in the image worker (worker.js) so the window
// keeps answering; work.js runs the same function in the window where a
// worker cannot be started. Every job is handed buffers and hands buffers back.
//
//   { op: 'wand', width, height, rgba, x, y, tolerance, contiguous }   → { mask }
//   { op: 'colour', width, height, rgba, colour, tolerance }          → { mask }
//   { op: 'reshape', width, height, mask, how, amount }               → { mask }
//   { op: 'outline', width, height, mask }                             → { segments }
//   { op: 'fill', width, height, rgba, region, colour, opacity }       → { rgba }
//   { op: 'gradient', width, height, rgba, region, gradient }          → { rgba }
//   { op: 'heal', width, height, src, dst, coverage, radius }          → { rgba }
import { magicWand, colourSelection, featherSelection, growSelection, shrinkSelection, selectionOutline } from '@rutba/imaging/select';
import { fillRegion, fillGradient, healPatch } from '@rutba/imaging/paint';

const u8 = (b) => (b ? new Uint8Array(b) : null);
const rgbaOf = (b) => new Uint8ClampedArray(b);

/** Run one job: { result, transfer }. */
export function runTool(job) {
  const { width: W, height: H } = job;
  const sel = (data) => ({ width: W, height: H, data });
  const maskOut = (s) => ({ result: { mask: s.data.buffer }, transfer: [s.data.buffer] });
  const pixOut = (d) => ({ result: { rgba: d.buffer }, transfer: [d.buffer] });
  switch (job.op) {
    case 'wand':
      return maskOut(magicWand(rgbaOf(job.rgba), W, H, job.x, job.y, { tolerance: job.tolerance, contiguous: job.contiguous }));
    case 'colour':
      return maskOut(colourSelection(rgbaOf(job.rgba), W, H, job.colour, { tolerance: job.tolerance }));
    case 'reshape': {
      const s = sel(u8(job.mask));
      const out = job.how === 'feather' ? featherSelection(s, job.amount) : job.how === 'grow' ? growSelection(s, job.amount) : shrinkSelection(s, job.amount);
      // A reshape that changed nothing answers a copy, so the buffer handed in can be given up.
      return maskOut(out === s ? sel(new Uint8Array(s.data)) : out);
    }
    case 'outline': {
      const segs = new Float32Array(selectionOutline(sel(u8(job.mask))));
      return { result: { segments: segs.buffer }, transfer: [segs.buffer] };
    }
    case 'fill':
      return pixOut(fillRegion(rgbaOf(job.rgba), W, H, u8(job.region), job.colour, { opacity: job.opacity ?? 1 }));
    case 'gradient':
      return pixOut(fillGradient(rgbaOf(job.rgba), W, H, { ...job.gradient, region: u8(job.region) }));
    case 'heal':
      return pixOut(healPatch(rgbaOf(job.src), rgbaOf(job.dst), W, H, { radius: job.radius, coverage: u8(job.coverage) }));
    default:
      throw new Error(`unknown tool job ${job.op}`);
  }
}
