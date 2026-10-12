// Picture Format → Edit in Image (docs/STUDIO.md I16), for Documents,
// Worksheets and Presentations alike: the picked picture's bytes are asked
// of the document, written to a scratch file and opened in an Image window
// that knows where it came from; when that window saves, the main process
// hands the edited picture to this window only (`picture:returned`), and
// the app puts it back where it was through its own edit, so Undo and Save
// see it as any other change.

import { t } from '@rutba/office-ui';

const EDITABLE = /^image\/(png|jpeg|gif|bmp|webp)$/;
const EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp' };

/** Pictures out with Image, by token: what puts each one back. */
const pending = new Map();
let listening = null;

function listen(shell) {
  if (listening || !shell?.on) return;
  listening = shell.on('picture:returned', async (back) => {
    const job = pending.get(back?.token);
    if (!job) return;
    try {
      await job.onBack(back);
    } catch (err) {
      job.toast?.(String(err?.message || err), { tone: 'bad', ms: 6000 });
    }
  });
}

const fromBase64 = (b64) => {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
};

/**
 * Open a picture of this document in Image. `sourceOp` is the document op
 * that hands its bytes out (`pictureSource`); `onBack({ bytes, contentType,
 * width, height })` puts the edited picture back. Answers false when the
 * picture is in a form Image cannot edit (a metafile, an SVG), having said so.
 * `extra` is what this app sends with every apply (a slide, a width).
 */
export async function editInImage({ shell, docId, sourceOp, name, onBack, toast, extra = {} }) {
  const got = await shell.doc.apply({ ...extra, id: docId, ops: [sourceOp] });
  const src = JSON.parse(got?.opResult || 'null');
  if (!src?.data) throw new Error(t('That picture\'s file is missing'));
  if (!EDITABLE.test(src.contentType)) {
    toast?.(t('This picture is kept in a form Image cannot edit (such as a drawing of lines). Change it to a PNG or a JPEG first.'), { ms: 6000 });
    return false;
  }
  const { path } = await shell.fs.temp({ ext: EXT[src.contentType], bytes: fromBase64(src.data) });
  const token = `p${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  pending.set(token, { onBack, toast });
  listen(shell);
  await shell.win.create({ app: 'image', file: path, query: { editFor: token, editName: name || t('the document'), editType: src.contentType } });
  return true;
}
