/**
 * How one sparkline is drawn, worked out from its group's settings and its
 * numbers: where the bottom and the top of its box sit on the value scale,
 * and which points are picked out in which colour. Kept apart from the
 * window so it can be checked without one.
 */

const isNumber = (v) => typeof v === 'number' && Number.isFinite(v);

/**
 * @param {object} group         a group as `Workbook.sparklineGroups` reads it
 * @param {Array<number|null>} values  this sparkline's cells, a blank as null
 * @param {Array<number|null>|null} groupValues  every cell of the group, when
 *   the group shares one axis
 * @returns {{ floor: number, ceil: number, points: Array<{ i: number, kind: string, colour: string }> }}
 */
export function sparklineShape(group, values, groupValues = null) {
  const nums = values.filter(isNumber);
  if (!nums.length) return { floor: 0, ceil: 0, points: [] };
  const shared = groupValues ? groupValues.filter(isNumber) : [];
  const whole = shared.length ? shared : nums;
  // Each end by itself: this sparkline's own, the group's, or a figure typed in.
  let min = group.minType === 'group' ? Math.min(...whole) : Math.min(...nums);
  let max = group.maxType === 'group' ? Math.max(...whole) : Math.max(...nums);
  if (group.minType === 'custom' && isNumber(group.manualMin)) min = group.manualMin;
  if (group.maxType === 'custom' && isNumber(group.manualMax)) max = group.manualMax;
  if (max < min) max = min;
  let floor = min;
  let ceil = max;
  if (group.type === 'column') {
    // A column stands on nought when nought is in sight. When every number
    // is above it (or below), the shortest column keeps a little height, so
    // 21 beside 24 reads as a rise and not as nothing beside something.
    const room = max > min ? (max - min) * 0.15 : Math.abs(max) || 1;
    if (min > 0 && group.minType !== 'custom') floor = Math.max(0, min - room);
    if (max < 0 && group.maxType !== 'custom') ceil = Math.min(0, max + room);
  }
  const marks = group.marks || {};
  const colours = group.colours || {};
  const points = [];
  const high = Math.max(...nums);
  const low = Math.min(...nums);
  let firstAt = -1;
  let lastAt = -1;
  values.forEach((v, i) => { if (isNumber(v)) { if (firstAt < 0) firstAt = i; lastAt = i; } });
  values.forEach((v, i) => {
    if (!isNumber(v)) return;
    // One colour for a point that is several things at once, the rarer first.
    const kind = marks.high && v === high ? 'high'
      : marks.low && v === low ? 'low'
        : marks.first && i === firstAt ? 'first'
          : marks.last && i === lastAt ? 'last'
            : marks.negative && v < 0 ? 'negative'
              : marks.markers && group.type !== 'column' ? 'markers' : null;
    if (kind) points.push({ i, kind, colour: colours[kind] || 'D00000' });
  });
  return { floor, ceil, points };
}
