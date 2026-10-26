// A turned layer resized from a corner, as Image's stage and Video's resize it:
// the corner opposite the grip stays where it stands turned.

import { layerBounds, applyLayerPatches } from '@rutba/studio/host';
import { framePoint } from '@rutba/office-ui/turned-box';

const measure = () => document.createElement('canvas').getContext('2d');

/**
 * A resize `patch` worked out as if the layer stood upright, moved so the
 * corner opposite the grip `kind` stays where it stands turned.
 */
export function keepTurnedCorner(plan, layer, start, kind, patch) {
  const rot = layer.rot || 0;
  if (!patch || !rot || !start) return patch;
  // Where the engine's patch puts it: the patch applied to a copy of the plan, as the next frame will.
  const copy = { ...plan, layers: plan.layers.map((l) => (l.id === layer.id ? { ...l } : l)) };
  applyLayerPatches(copy, [{ ...patch, id: layer.id }]);
  const b1 = layerBounds(measure(), copy, copy.layers.find((l) => l.id === layer.id));
  if (!b1) return patch;
  const opposite = [kind.includes('w') ? 1 : 0, kind.includes('n') ? 1 : 0];
  const want = framePoint(start, rot, ...opposite);
  const got = framePoint(b1, rot, ...opposite);
  const fx = (patch.fx ?? layer.fx ?? 0.5) + (want.x - got.x) / plan.W;
  const fy = (patch.fy ?? layer.fy ?? 0.5) + (want.y - got.y) / plan.H;
  return { ...patch, fx: +fx.toFixed(4), fy: +fy.toFixed(4) };
}
