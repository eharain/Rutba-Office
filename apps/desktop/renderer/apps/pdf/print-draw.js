// The pages a print needs, drawn one at a time before the system's dialog
// opens. A long file takes a while (a few pages a second), so the drawing
// says how far it has got and stops when the person asks: what was drawn so
// far is let go, and nothing is sent to the printer.

/**
 * Draws every page of a list, in order.
 *
 * @param {object[]} list                       the pages, as the window keeps them
 * @param {(page: object, index: number) => Promise<object>} draw   draws one, answering its picture
 * @param {object} [o]
 * @param {(n: number, count: number) => void} [o.onPage]   told before each page is drawn, counted from 1
 * @param {() => boolean} [o.stopped]           asked before each page, and once more after the last
 * @param {(picture: object) => void} [o.drop]  lets go of a picture that will not be printed
 * @returns {Promise<object[]|null>} the pictures, or null when it was stopped
 */
export async function drawForPrint(list, draw, { onPage = () => {}, stopped = () => false, drop = () => {} } = {}) {
  const pictures = [];
  const letGo = () => { for (const p of pictures.splice(0)) drop(p); };
  try {
    for (let i = 0; i < list.length; i += 1) {
      if (stopped()) { letGo(); return null; }
      onPage(i + 1, list.length);
      pictures.push(await draw(list[i], i));
    }
    if (stopped()) { letGo(); return null; }
    return pictures;
  } catch (err) {
    letGo();
    throw err;
  }
}

/** Whether a print the system answered with a reason was the person closing its dialog: nothing to tell. */
export function printCancelled(answer) {
  return !answer?.ok && (!answer?.reason || /cancel/i.test(String(answer.reason)));
}
