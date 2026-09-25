// Review → Check Accessibility: the rules.
//
// Each app's reader (word.js, sheet.js, deck.js) describes its
// document in one neutral shape — the objects that need alternative text,
// the headings, the tables, the text and what it sits on, the links, the
// sheets, the slides — and the rules here turn that description into
// findings. The rules know nothing about XML and the readers know nothing
// about what makes a finding: the same "Missing alternative text" holds a
// picture in a letter, a chart on a sheet and a logo on a slide to one bar.
//
// A finding is an Error (the content is hard or impossible to read for
// someone with a disability), a Warning (hard for most) or a Tip (could be
// organised better) — Office's three tiers and its own rule names, so a
// person who has used Office's checker recognises every line. Each finding
// says where it is, in the app's own coordinates, and carries the one-click
// fix when there is one.

import { contrastRatio, requiredRatio, readableOn } from './colour.js';

export const SEVERITIES = [
  { id: 'error', label: 'Errors' },
  { id: 'warning', label: 'Warnings' },
  { id: 'tip', label: 'Tips' },
];

/**
 * The rules, in the order the pane lists them within a tier. `why` and `how`
 * are the pane's "Additional information" — Office writes the same two
 * paragraphs under every finding.
 */
export const RULES = {
  altText: {
    severity: 'error',
    title: 'Missing alternative text',
    why: 'Alternative text helps people who use screen readers understand what is shown in pictures, charts and other objects.',
    how: 'Select the object and use Edit Alt Text to describe it in a sentence or two. If it is only there to look good, mark it as decorative instead.',
  },
  tableHeader: {
    severity: 'error',
    title: 'Missing table header',
    why: 'A header row tells a screen reader what each column holds, so a person hears the column name with every cell rather than a list of numbers.',
    how: 'Make the first row of the table a header row, and give each column a short name in it.',
  },
  slideTitle: {
    severity: 'error',
    title: 'Missing slide title',
    why: 'Slide titles are how people who cannot see the screen find their way around a presentation, and how they choose a slide to jump to.',
    how: 'Type a title in the slide\'s title placeholder. A slide with no title placeholder can take a layout that has one.',
  },
  contrast: {
    severity: 'warning',
    title: 'Hard-to-read text contrast',
    why: 'Text that is too close in colour to what is behind it is hard to read for people with low vision or colour blindness, and for everyone on a bright screen.',
    how: 'Use a darker text colour on a light background, or a lighter one on a dark background. Text needs a contrast of 4.5:1, large text 3:1.',
  },
  headingOrder: {
    severity: 'warning',
    title: 'Skipped heading level',
    why: 'Headings in order — a Heading 2 under a Heading 1, never straight to a Heading 3 — let a screen reader describe the document\'s outline as it really is.',
    how: 'Change the heading to the next level down from the one above it.',
  },
  emptyHeading: {
    severity: 'warning',
    title: 'Empty heading',
    why: 'An empty paragraph in a heading style is announced as a heading with nothing in it, and shows as a blank line in the navigation pane and the table of contents.',
    how: 'Type the heading, or give the paragraph the Normal style.',
  },
  mergedCells: {
    severity: 'warning',
    title: 'Merged cells',
    why: 'Screen readers keep their place in a table by counting cells. A merged cell throws the count out, and what is read next is no longer what is under the heading.',
    how: 'Unmerge the cells, and use a simple table structure with one piece of information in each cell.',
  },
  linkText: {
    severity: 'warning',
    title: 'Unclear hyperlink text',
    why: 'People who use screen readers often listen to a list of a document\'s links on their own. "Click here" or a long web address says nothing about where the link goes.',
    how: 'Change the link\'s words to say where it goes — "the 2026 price list", not "click here".',
  },
  duplicateTitle: {
    severity: 'warning',
    title: 'Duplicate slide title',
    why: 'Two slides with the same title are two identical entries in the list of slides a screen reader offers, and a person cannot tell which is which.',
    how: 'Give each slide a title of its own — "Results (continued)" for a second slide on the same subject.',
  },
  readingOrder: {
    severity: 'warning',
    title: 'Check reading order',
    why: 'A screen reader reads a slide in the order its objects are layered, not the order they appear. When the two differ, the slide is read out of order.',
    how: 'Open the Layers pane and put the objects in the order they should be read — the first to be read at the bottom of the list.',
  },
  titleMissing: {
    severity: 'warning',
    title: 'Document title missing',
    why: 'The title in a document\'s properties is what a screen reader announces when the file opens, and what search results show.',
    how: 'Set a title for the document.',
  },
  blankLines: {
    severity: 'tip',
    title: 'Repeated blank characters',
    why: 'Empty paragraphs used to make space are read out as "blank" again and again, and move when the text above them changes.',
    how: 'Remove the empty paragraphs and use spacing before and after the paragraphs instead.',
  },
  defaultSheetName: {
    severity: 'tip',
    title: 'Default sheet names',
    why: 'A sheet called "Sheet1" says nothing. A screen reader announces sheet names as a person moves between them.',
    how: 'Rename the sheet to say what is on it.',
  },
};

/** Words a link should never be made of on their own. */
const VAGUE_LINK = /^(click|click here|here|link|this link|this|more|read more|learn more|go|go here|details|info|more info|website|page|this page)$/i;
// A description that is only the picture's file name says nothing: "IMG_2041.png".
const FILE_NAME = /^[^\s\\/]+\.(?:png|jpe?g|gif|bmp|webp|tiff?|svg|emf|wmf|heic)$/i;
const BARE_URL = /^(?:https?:\/\/|ftp:\/\/|www\.)\S+$/i;

export function unclearLinkText(text) {
  const t = String(text || '').replace(/\s+/g, ' ').trim().replace(/[.!:]+$/, '');
  if (!t) return false;
  return VAGUE_LINK.test(t) || BARE_URL.test(t);
}

/** "Sheet1", "Sheet 2", "Feuil3" is not ours to judge — Excel's own default only. */
export function isDefaultSheetName(name) {
  return /^Sheet\s?\d+$/i.test(String(name || '').trim());
}

/**
 * A slide's objects in the order a person reads them: top to bottom, then
 * left to right among the objects on the same band — two boxes whose tops
 * are within a quarter of the shorter one's height count as one row.
 */
export function visualOrder(shapes) {
  const list = shapes.filter((s) => s && Number.isFinite(s.y) && Number.isFinite(s.x));
  return [...list].sort((a, b) => {
    const band = Math.max(8, Math.min(a.h || 0, b.h || 0) / 4);
    if (Math.abs(a.y - b.y) > band) return a.y - b.y;
    return a.x - b.x;
  });
}

/**
 * Run every rule over a document's description. Returns the findings in the
 * pane's order — tier, then rule, then position — and the verdict the status
 * bar shows.
 */
export function checkAccessibility(model = {}) {
  const issues = [];
  const add = (rule, item) => issues.push({ rule, severity: RULES[rule].severity, ...item });

  // Missing alternative text — pictures, charts, shapes without words, groups.
  for (const o of model.objects || []) {
    const alt = String(o.alt || '').trim();
    if (o.decorative || (alt && !FILE_NAME.test(alt))) continue;
    add('altText', { key: `alt:${o.key}`, label: o.label, where: o.where, fixes: [{ kind: 'altText', label: 'Add a description', target: o.target }, { kind: 'decorative', label: 'Mark as decorative', target: o.target }] });
  }

  // Tables: a header row, and no merged cells.
  for (const t of model.tables || []) {
    if (!t.hasHeader) add('tableHeader', { key: `th:${t.key}`, label: t.label, where: t.where, fixes: t.headerFix ? [{ kind: 'tableHeader', label: t.headerFix, target: t.target }] : [] });
    if (t.merged) add('mergedCells', { key: `tm:${t.key}`, label: t.label, where: t.where, fixes: [] });
  }
  for (const m of model.merges || []) {
    add('mergedCells', { key: `mc:${m.key}`, label: m.label, where: m.where, fixes: [{ kind: 'unmerge', label: 'Unmerge', target: m.target }] });
  }

  // Headings in order, and none empty.
  let previous = 0;
  for (const h of model.headings || []) {
    if (!String(h.text || '').trim()) {
      add('emptyHeading', { key: `he:${h.key}`, label: h.label || `Empty ${h.styleName || 'heading'}`, where: h.where, fixes: [{ kind: 'normalStyle', label: 'Change to Normal', target: h.target }] });
      continue;
    }
    if (h.level > previous + 1) {
      const level = previous + 1;
      add('headingOrder', { key: `ho:${h.key}`, label: h.label, where: h.where, detail: `Heading ${h.level} after ${previous ? `Heading ${previous}` : 'no heading'}`, fixes: [{ kind: 'headingLevel', label: `Change to Heading ${level}`, level, target: h.target }] });
      previous = level;
    } else {
      previous = h.level;
    }
  }

  // Text contrast — one finding per paragraph, cell or shape, whatever the
  // number of runs in it that fall short.
  const contrastSeen = new Set();
  for (const t of model.texts || []) {
    if (!t.fg || !t.bg || contrastSeen.has(t.key)) continue;
    const ratio = contrastRatio(t.fg, t.bg);
    if (ratio == null || ratio >= requiredRatio(t.sizePt, t.bold)) continue;
    contrastSeen.add(t.key);
    const colour = readableOn(t.bg);
    add('contrast', {
      key: `co:${t.key}`, label: t.label, where: t.where,
      detail: `${ratio.toFixed(1)}:1 — ${t.fg} on ${t.bg}`,
      fixes: t.target ? [{ kind: 'textColour', label: colour === '#FFFFFF' ? 'Use white text' : 'Use black text', colour, target: t.target }] : [],
    });
  }

  // Links whose words say nothing about where they go.
  for (const l of model.links || []) {
    if (!unclearLinkText(l.text)) continue;
    add('linkText', { key: `lk:${l.key}`, label: l.label || `"${String(l.text).trim()}"`, where: l.where, detail: l.url || '', fixes: [] });
  }

  // Empty paragraphs used as spacing.
  for (const b of model.blankRuns || []) {
    add('blankLines', { key: `bl:${b.key}`, label: b.label, where: b.where, fixes: [{ kind: 'removeBlanks', label: 'Remove the empty paragraphs', target: b.target }] });
  }

  // Sheet names.
  for (const s of model.sheets || []) {
    if (!isDefaultSheetName(s.name)) continue;
    add('defaultSheetName', { key: `sn:${s.name}`, label: s.name, where: s.where, fixes: [{ kind: 'renameSheet', label: 'Rename the sheet', target: s.name }] });
  }

  // Slides: a title each, no two the same, and a reading order that follows the eye.
  const titles = new Map();
  for (const s of model.slides || []) {
    const title = String(s.title || '').replace(/\s+/g, ' ').trim();
    if (!title) {
      add('slideTitle', { key: `st:${s.index}`, label: `Slide ${s.index + 1}`, where: s.where, fixes: s.titleShape != null ? [{ kind: 'slideTitle', label: 'Type a slide title', target: { slide: s.index, shape: s.titleShape } }] : [] });
    } else {
      const key = title.toLowerCase();
      if (!titles.has(key)) titles.set(key, []);
      titles.get(key).push(s);
    }
    const shapes = (s.shapes || []).filter((x) => !x.hidden);
    if (shapes.length > 1) {
      const eye = visualOrder(shapes).map((x) => String(x.id));
      const layered = shapes.map((x) => String(x.id));
      if (eye.join('|') !== layered.join('|')) {
        add('readingOrder', { key: `ro:${s.index}`, label: `Slide ${s.index + 1}`, where: s.where, detail: 'The objects are layered in a different order from where they sit.', fixes: [{ kind: 'layers', label: 'Open the Layers pane', target: { slide: s.index } }] });
      }
    }
  }
  for (const list of titles.values()) {
    if (list.length < 2) continue;
    for (const s of list) {
      add('duplicateTitle', { key: `dt:${s.index}`, label: `Slide ${s.index + 1}: ${String(s.title).trim()}`, where: s.where, fixes: s.titleShape != null ? [{ kind: 'slideTitle', label: 'Change the slide title', target: { slide: s.index, shape: s.titleShape } }] : [] });
    }
  }

  // The document's title.
  if (model.title !== undefined && !String(model.title || '').trim()) {
    add('titleMissing', { key: 'title', label: 'Document properties', where: null, fixes: [{ kind: 'setTitle', label: 'Set a title', suggestion: model.titleSuggestion || '' }] });
  }

  const tierOrder = { error: 0, warning: 1, tip: 2 };
  const ruleOrder = Object.keys(RULES);
  const sorted = issues
    .map((issue, i) => ({ issue, i }))
    .sort((a, b) => tierOrder[a.issue.severity] - tierOrder[b.issue.severity] || ruleOrder.indexOf(a.issue.rule) - ruleOrder.indexOf(b.issue.rule) || a.i - b.i)
    .map((x) => x.issue);
  const counts = { error: 0, warning: 0, tip: 0 };
  for (const issue of sorted) counts[issue.severity] += 1;
  return { issues: sorted, counts, verdict: counts.error || counts.warning ? 'investigate' : 'good' };
}

/** Findings grouped the way the pane lists them: tier → rule → items. */
export function groupIssues(issues = []) {
  return SEVERITIES.map((tier) => {
    const inTier = issues.filter((i) => i.severity === tier.id);
    const rules = [];
    for (const issue of inTier) {
      let group = rules.find((r) => r.rule === issue.rule);
      if (!group) rules.push((group = { rule: issue.rule, title: RULES[issue.rule]?.title || issue.rule, items: [] }));
      group.items.push(issue);
    }
    return { ...tier, count: inTier.length, rules };
  }).filter((t) => t.count);
}
