// What each kind of clip is called and drawn with, said once for the
// timeline, the properties pane and the undo list. The engine keeps its kinds
// in English (packages/studio); each is written out here through t() so the
// catalogue carries it in every language.

import { t } from '@rutba/office-ui';

/** The icon a clip of each kind wears. */
export const KIND_ICONS = {
  video: 'video', image: 'picture', sound: 'volume', text: 'textbox', shape: 'shape', icon: 'star', wordart: 'textbox', pointer: 'pointer',
  textblock: 'list', chart: 'chart', adjust: 'sliders', qr: 'qrcode', gradient: 'gradient', progress: 'minus', edges: 'contrast', outro: 'textbox',
};

/** The kinds that stand in a box of their own on the frame, which the grips on the preview move and resize. */
export const BOXED_KINDS = ['textblock', 'chart', 'qr'];

/** The kinds that lie over the whole frame: picked on the timeline, never on the preview. */
export const OVERLAY_KINDS = ['adjust', 'gradient', 'progress', 'edges', 'outro'];

/** The chart kinds, by the studio's keys. */
export const CHART_NAMES = () => ({
  bar: t('image.words.bars'), column: t('common.columns'), line: t('common.line'), area: t('common.area'), pie: t('common.pie'), donut: t('image.words.donut'),
});

/** The colour looks an adjustment starts from, by the studio's preset keys. */
export const ADJUST_NAMES = () => ({
  none: t('common.none'), punch: t('image.words.punch'), matte: t('image.words.matte'), warm: t('image.words.warm'), cool: t('image.words.cool'), mono: t('image.words.mono'), fade: t('common.fade'), noir: t('image.words.noir'),
});

/** The ways a paragraph's lines are marked. */
export const LIST_NAMES = () => ({
  none: t('common.paragraph'), bullet: t('mail.compose.bulleted-list'), dash: t('video.layer-names.dashed-list'), number: t('mail.compose.numbered-list'), check: t('video.layer-names.checklist'),
});

/** The overlays, by the keys of the studio's OVERLAYS. */
export const OVERLAY_NAMES = () => ({
  qr: t('common.qr-code'), progress: t('video.layer-names.progress-bar'), gradient: t('video.layer-names.bottom-shade'), edges: t('video.layer-names.fade-out'), outro: t('video.layer-names.outro-card'),
});

/** What a clip of each kind is called before it has words or a name of its own. */
export const kindNames = () => ({
  video: t('video.video-clip'), image: t('common.picture'), sound: t('video.sound-clip'), text: t('common.title'), shape: t('common.shape'), icon: t('common.icon'), wordart: t('video.word-art'),
  caption: t('common.caption'), pointer: t('video.pointer'),
  textblock: t('common.paragraph'), chart: t('common.chart'), adjust: t('image.adjustment'), ...OVERLAY_NAMES(),
});

/** The first line of a clip's words, for a label. */
const firstLine = (text) => String(text || '').split('\n').map((s) => s.trim()).find(Boolean) || '';

/**
 * A clip's name as the timeline and the properties pane show it: its words,
 * else what it is. Media clips are named by their file, which the caller has.
 */
export function layerLabel(clip) {
  const kinds = kindNames();
  switch (clip?.type) {
    case 'textblock': return firstLine(clip.text) || LIST_NAMES()[clip.list || 'none'] || kinds.textblock;
    case 'chart': return t('image.words.kind-chart', { kind: CHART_NAMES()[clip.chart || 'bar'] || '' }).trim();
    case 'adjust': return t('image.words.look-adjustment', { look: ADJUST_NAMES()[clip.preset || 'none'] || '' }).trim();
    case 'qr': return clip.data ? `${kinds.qr}: ${clip.data}` : kinds.qr;
    case 'outro': return firstLine(clip.text) || kinds.outro;
    case 'gradient':
    case 'progress':
    case 'edges':
      return kinds[clip.type];
    default: return '';
  }
}
