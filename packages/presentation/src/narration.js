/**
 * Record → narration: a slide's recorded voice as PowerPoint keeps it — an
 * audio shape ("Audio Recording") the show hides, and in the slide's timing
 * a media node for it and a call that plays it from the start as the slide
 * comes in (the main sequence's first group, begun by the slide itself).
 * Nothing here touches a package: this edits one slide's XML.
 */
import { timingRange, insertTiming } from './timing.js';

const MAIN_SEQ_CONDS = '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>';

/** The narration's name: what marks a slide's recorded voice among its sounds. */
export const NARRATION_NAME = 'Audio Recording';
export const isNarration = (shape) => Boolean(shape?.media?.kind === 'audio' && /^Audio Recording\b/.test(shape.name || ''));

/**
 * The slide with `spid` (its narration's shape) played from the start as
 * the slide comes in, for `durationMs` — the call put first in the main
 * sequence (in the group the slide begins, or in a new one), and the
 * shape's media node beside the main sequence, hidden while it is not
 * playing, as PowerPoint writes a recorded narration.
 */
export function withNarration(slideXml, spid, durationMs) {
  let top = Math.max(0, ...[...slideXml.matchAll(/<p:cTn\b[^>]*?\sid="(\d+)"/g)].map((m) => Number(m[1])));
  const id = () => ++top;
  const call = () => `<p:par><p:cTn id="${id()}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
    + `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${id()}" dur="${Math.max(1, Math.round(durationMs))}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`;
  const media = () => `<p:audio><p:cMediaNode vol="80000" showWhenStopped="0"><p:cTn id="${id()}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>`
    + `<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`;
  const autoGroup = (mainId) => `<p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${mainId}"/></p:cond></p:stCondLst>`
    + `<p:childTnLst><p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${call()}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;

  const range = timingRange(slideXml);
  if (!range) {
    const root = id();
    const main = id();
    const timing = `<p:timing><p:tnLst><p:par><p:cTn id="${root}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
      + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${main}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${autoGroup(main)}</p:childTnLst></p:cTn>${MAIN_SEQ_CONDS}</p:seq>`
      + `${media()}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    return insertTiming(slideXml, timing);
  }
  let t = slideXml.slice(range.start, range.end);
  const mainOpen = /<p:cTn\b[^>]*\bnodeType="mainSeq"[^>]*>/.exec(t);
  if (mainOpen) {
    const mainId = /\sid="(\d+)"/.exec(mainOpen[0])[1];
    const kids = t.indexOf('<p:childTnLst>', mainOpen.index + mainOpen[0].length);
    if (kids < 0) {
      // A main sequence with nothing in it yet.
      const at = mainOpen.index + mainOpen[0].length;
      t = t.slice(0, at) + `<p:childTnLst>${autoGroup(mainId)}</p:childTnLst>` + t.slice(at);
    } else {
      const inner = kids + '<p:childTnLst>'.length;
      // The group the slide itself begins, when there is one: the call goes first in it.
      const first = /^<p:par><p:cTn\b[^>]*>(<p:stCondLst>[\s\S]*?<\/p:stCondLst>)<p:childTnLst><p:par><p:cTn\b[^>]*><p:stCondLst>[\s\S]*?<\/p:stCondLst><p:childTnLst>/.exec(t.slice(inner));
      t = first && /evt="onBegin"/.test(first[1])
        ? t.slice(0, inner + first[0].length) + call() + t.slice(inner + first[0].length)
        : t.slice(0, inner) + autoGroup(mainId) + t.slice(inner);
    }
  } else {
    const rootOpen = /<p:cTn\b[^>]*\bnodeType="tmRoot"[^>]*>/.exec(t);
    if (!rootOpen) return slideXml;
    const main = id();
    const seq = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${main}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${autoGroup(main)}</p:childTnLst></p:cTn>${MAIN_SEQ_CONDS}</p:seq>`;
    const kids = t.indexOf('<p:childTnLst>', rootOpen.index);
    t = kids >= 0
      ? t.slice(0, kids + '<p:childTnLst>'.length) + seq + t.slice(kids + '<p:childTnLst>'.length)
      : t.slice(0, rootOpen.index + rootOpen[0].length) + `<p:childTnLst>${seq}</p:childTnLst>` + t.slice(rootOpen.index + rootOpen[0].length);
  }
  // The media node, last under the root.
  const end = t.lastIndexOf('</p:childTnLst></p:cTn></p:par></p:tnLst>');
  if (end >= 0) t = t.slice(0, end) + media() + t.slice(end);
  return slideXml.slice(0, range.start) + t + slideXml.slice(range.end);
}

/** The slide without `spid`'s media node — what a cleared narration leaves behind once its call is gone. */
export function withoutMediaNode(slideXml, spid) {
  return slideXml.replace(/<p:(audio|video)>(?:(?!<\/p:\1>)[\s\S])*?<\/p:\1>/g, (node) => (new RegExp(`<p:spTgt spid="${spid}"`).test(node) ? '' : node));
}
