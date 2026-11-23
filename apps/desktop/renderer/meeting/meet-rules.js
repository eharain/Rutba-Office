// What a meeting, as the introducer shows it, lets a person do: read by the
// room (meet-room.js) and by Talk's own screens, which take these from
// apps/talk/meet-lines.js under the same names. No words and no drawing.

/**
 * Why someone looking at a meeting cannot come into it now, from what they
 * were shown of it: 'full', 'locked', or null when they can. Someone already
 * in it on another computer can always bring it here.
 */
export function cannotJoin(meeting, meId) {
  const people = meeting?.people || [];
  if (people.some((p) => p.id === meId)) return null;
  if (meeting?.locked && meeting?.starter?.id !== meId) return 'locked';
  return people.length >= (meeting?.most || 4) ? 'full' : null;
}

/** Whether a meeting, as the introducer shows it, is held on a Talk Server. */
export const isCarried = (meeting) => meeting?.carried === 'server';
