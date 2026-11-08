// The bar of icons at the top of the Live panel, as plain functions (owner,
// 8 Nov: "the icons on top should let user enable the streaming destinations").
//
// What each stream's icon says, what Wide and Tall do, whether Go live can be
// pressed, and what the bar's words are. No React and no window, so the rules
// are tested in Node.

/** @typedef {{ id: string, shape?: string, hasKey?: boolean, name?: string, platform?: string }} Stream */

export const shapeOf = (d) => (d?.shape === 'tall' ? 'tall' : 'wide');

/** Stream ids that are on, in list order, with those removed or without a key dropped. */
export function reconcileOn(on, streams) {
  const wanted = new Set(on || []);
  return streams.filter((d) => wanted.has(d.id) && d.hasKey).map((d) => d.id);
}

/** A stream's icon clicked: on or off. A stream with no key cannot be turned on. */
export function toggleStream(on, streams, id) {
  const d = streams.find((x) => x.id === id);
  if (!d?.hasKey) return reconcileOn(on, streams);
  const next = new Set(on);
  if (next.has(id)) next.delete(id); else next.add(id);
  return reconcileOn([...next], streams);
}

/**
 * Wide or Tall clicked: the whole group off, remembering which were on, and
 * clicked again they come back. `memory` is what the last call answered.
 */
export function toggleGroup(on, streams, shape, memory = {}) {
  const group = streams.filter((d) => shapeOf(d) === shape && d.hasKey).map((d) => d.id);
  const inGroup = new Set(group);
  const nowOn = (on || []).filter((id) => inGroup.has(id));
  if (nowOn.length) {
    return { on: reconcileOn((on || []).filter((id) => !inGroup.has(id)), streams), memory: { ...memory, [shape]: nowOn } };
  }
  return { on: reconcileOn([...(on || []), ...(memory[shape] || [])], streams), memory: { ...memory, [shape]: [] } };
}

/**
 * The whole bar's state.
 * `status` is what the broadcast says of each stream that is going;
 * `open` is whether a broadcast is open.
 * Tile states: nokey, off, on, live, retrying, refused.
 * Air: ready, connecting, live, offline.
 */
export function barModel({ streams = [], on = [], status = [], open = false } = {}) {
  const ids = new Set(reconcileOn(on, streams));
  const byId = new Map(status.map((s) => [s.id, s]));
  const tiles = streams.map((d) => {
    const isOn = ids.has(d.id);
    const s = byId.get(d.id);
    let state = !d.hasKey ? 'nokey' : isOn ? 'on' : 'off';
    if (isOn && open && s) {
      if (s.state === 'live') state = 'live';
      else if (s.state === 'reconnecting') state = 'retrying';
      else if (s.state === 'error') state = 'refused';
    }
    return { id: d.id, stream: d, shape: shapeOf(d), on: isOn, state, detail: s?.detail || '' };
  });
  const group = (shape) => {
    const list = tiles.filter((x) => x.shape === shape);
    return { shape, tiles: list, on: list.some((x) => x.on) };
  };
  const onIds = tiles.filter((x) => x.on).map((x) => x.id);
  const liveCount = tiles.filter((x) => x.state === 'live').length;
  const trying = tiles.some((x) => x.on && (x.state === 'on' || x.state === 'retrying'));
  const air = !open ? 'ready' : liveCount ? 'live' : trying ? 'connecting' : 'offline';
  return { tiles, wide: group('wide'), tall: group('tall'), onIds, count: onIds.length, liveCount, canGo: onIds.length > 0, air };
}

/** Whether a shape has any stream on: the encoders run only for the shapes that do. */
export function shapesOn(streams, on) {
  const ids = new Set(on);
  return ['wide', 'tall'].filter((shape) => streams.some((d) => shapeOf(d) === shape && d.hasKey && ids.has(d.id)));
}

/** The upload a set of streams needs, in Mbit/s: each stream its shape's picture and the sound, with room to spare. */
export function uploadMbps(tiles, kbpsOf, audioKbps = 128) {
  let kbps = 0;
  for (const x of tiles) if (x.on) kbps += (kbpsOf[x.shape] || 0) + audioKbps;
  return (kbps * 1.15) / 1000;
}

/** What Go live and the icons have to do when the choices change during a broadcast. */
export function onChanges(before, after) {
  const was = new Set(before);
  const now = new Set(after);
  return { start: after.filter((id) => !was.has(id)), stop: before.filter((id) => !now.has(id)) };
}
