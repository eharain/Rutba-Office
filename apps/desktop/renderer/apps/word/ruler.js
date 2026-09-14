// The ruler above the page and the grips on a table — the parts of Word a
// hand reaches for. The ruler shows the margins, the current paragraph's
// indents and tab stops and, with the caret in a table, its column edges;
// every one of them drags. The grips sit on a table's borders on the page
// itself. Both read what is DRAWN — a paragraph's computed style, a table's
// rectangles — rather than re-deriving it from the model, so a marker sits
// exactly where the words are; and both tell the engine once, on release,
// with a guide line following the hand until then.

import React from 'react';

const PX_PER_CM = 96 / 2.54;
const TWIPS_PER_PX = 15;
const TAB_TYPES = ['left', 'center', 'right', 'decimal'];
const MIN_COLUMN = 20; // px — the engine's own floor is a quarter centimetre
const MIN_ROW = 12;

const num = (v) => parseFloat(v) || 0;
const tw = (pxValue) => Math.round(pxValue * TWIPS_PER_PX);
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

/** The tab stops a style gives a block — what a paragraph inherits when it sets none. */
const styleStops = (block, styles) => (styles ? (styles[block?.style] ?? styles['*default*'])?.tabs : null) ?? [];

/**
 * A paragraph as drawn: the left edge of its lines after the first, its
 * first line, its right edge — in px from the page's left edge — and the
 * numbers the paragraph itself owns (its left indent and the first line's
 * offset from it), which are what the engine is told back.
 */
function measureParagraph(page, index) {
  const el = page?.querySelector(`.wd-block[data-block="${index}"]`);
  if (!el) return null;
  const cs = getComputedStyle(el);
  const r = el.getBoundingClientRect();
  const p = page.getBoundingClientRect();
  const paddingLeft = num(cs.paddingLeft);
  const textIndent = num(cs.textIndent);
  const left = r.left - p.left + paddingLeft;
  return {
    left,
    first: left + textIndent,
    right: r.right - p.left,
    edge: r.left - p.left - num(cs.marginLeft),
    leftPx: num(cs.marginLeft) + paddingLeft,
    indentPx: textIndent,
    rightPx: num(cs.marginRight),
  };
}

/**
 * A top-level table's parts on the page (one per page it spans): the x of
 * each column edge, by grid boundary, and the y of each row's bottom, by
 * row. The edges come from the cells when a row has one per column, and
 * from the grid's shares of the drawn width when every row merges.
 */
function measureTable(page, tableId, gridPx) {
  if (!page || !tableId || !gridPx?.length) return null;
  const p = page.getBoundingClientRect();
  const parts = [...page.querySelectorAll(`.wd-table[data-table="${tableId}"]`)];
  if (!parts.length) return null;
  const sum = gridPx.reduce((a, b) => a + b, 0);
  return parts.map((table) => {
    const t = table.getBoundingClientRect();
    const scale = sum > 0 ? t.width / sum : 1;
    const rows = [...table.querySelectorAll(':scope > tbody > tr')];
    const plain = rows.find((tr) => tr.children.length === gridPx.length && ![...tr.children].some((td) => td.colSpan > 1));
    const xs = plain
      ? [...plain.children].map((td, i) => ({ k: i + 1, x: td.getBoundingClientRect().right - p.left }))
      : gridPx.map((_, i) => ({ k: i + 1, x: t.left - p.left + gridPx.slice(0, i + 1).reduce((a, b) => a + b, 0) * scale }));
    const from = Number(table.dataset.rowFrom || 0);
    const ys = rows.map((tr, i) => {
      const r = tr.getBoundingClientRect();
      return { r: from + i, y: r.bottom - p.top, top: r.top - p.top };
    });
    return { left: t.left - p.left, top: t.top - p.top, width: t.width, height: t.height, scale, xs, ys };
  });
}

/** Re-measure whenever the page changes size (a zoom, a repagination). */
function useResize(page, measure) {
  React.useEffect(() => {
    const el = page.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    const ro = new ResizeObserver(() => measure());
    ro.observe(el);
    return () => ro.disconnect();
  }, [measure, page]);
}

/**
 * The ruler. Centimetres from the left margin, the margins shaded and
 * their edges draggable; for the paragraph at the caret the first-line
 * marker (top), the hanging marker (bottom) with the left-indent box under
 * it that moves both, and the right marker; its tab stops, which a click on
 * the ruler adds (the box at the far left picks the type), a drag moves and
 * a drag off the ruler removes; and, with the caret in a table, the
 * column edges. Everything writes through the callbacks, once, on release.
 */
export function Ruler({ section, page, model, at, tableId, gridPx, onParagraph, onMargin, onColumn }) {
  const W = section ? Math.round(section.widthPx) : 794;
  const ML = Math.round(section?.margins?.left ?? 96);
  const MR = Math.round(section?.margins?.right ?? 96);
  const ref = React.useRef(null);
  const [para, setPara] = React.useState(null);
  const [table, setTable] = React.useState(null);
  const [drag, setDrag] = React.useState(null);
  const [tabType, setTabType] = React.useState('left');

  const measure = React.useCallback(() => {
    setPara(measureParagraph(page.current, at));
    setTable(tableId ? measureTable(page.current, tableId, gridPx)?.[0] ?? null : null);
  }, [page, at, tableId, gridPx]);
  React.useLayoutEffect(() => {
    measure();
  }, [measure, model, section]);
  useResize(page, measure);

  const block = model?.blocks?.[at];
  const inherited = styleStops(block, model?.resolvedStyles);
  const stops = block ? block.tabs ?? inherited : [];

  // What the paragraph's stops become: its own list, or — none left — a
  // clear for each of the style's, so they do not come back.
  const writeStops = (list) => {
    const own = list.filter((s) => s.posPx > 0).map((s) => ({ align: s.align || 'left', posTwips: tw(s.posPx), ...(s.leader ? { leader: s.leader } : {}) }));
    const clears = own.length ? [] : inherited.map((s) => ({ align: 'clear', posTwips: tw(s.posPx) }));
    onParagraph({ tabs: [...own, ...clears] });
  };

  // Take hold of a marker: it follows the pointer between its bounds and
  // reports where it was let go. A tab stop dragged well above or below the
  // ruler is let go as "removed".
  const hold = (e, spec) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const x0 = e.clientX;
    const bar = ref.current.getBoundingClientRect();
    let x = spec.x;
    let off = false;
    setDrag({ kind: spec.kind, index: spec.index, x, off });
    const move = (ev) => {
      x = clamp(spec.x + ev.clientX - x0, spec.min, spec.max);
      off = spec.kind === 'tab' && (ev.clientY > bar.bottom + 24 || ev.clientY < bar.top - 24);
      setDrag({ kind: spec.kind, index: spec.index, x, off });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      const dx = Math.round(x - spec.x);
      if (off) spec.release(null, 0);
      else if (dx !== 0) spec.release(Math.round(x), dx);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  const pos = (kind, index, base) => (drag && drag.kind === kind && drag.index === index ? drag.x : base);

  // A click on the bare ruler, between the margins, adds a stop of the chosen type.
  const addStop = (e) => {
    if (e.button !== 0 || !block || !onParagraph) return;
    const x = e.clientX - ref.current.getBoundingClientRect().left;
    if (x <= ML || x >= W - MR) return;
    e.preventDefault();
    writeStops([...stops, { align: tabType, posPx: Math.round(x - ML) }]);
  };

  const ticks = [];
  for (let i = -Math.floor(ML / PX_PER_CM); i <= Math.floor((W - ML) / PX_PER_CM); i++) {
    const x = ML + i * PX_PER_CM;
    if (i !== 0 && x > 8 && x < W - 8) ticks.push(<span key={`n${i}`} className="wd-ruler-tick major" style={{ left: x }}>{Math.abs(i)}</span>);
    if (x + PX_PER_CM / 2 < W) ticks.push(<span key={`h${i}`} className="wd-ruler-tick half" style={{ left: x + PX_PER_CM / 2 }} />);
    for (const q of [0.25, 0.75]) if (x + q * PX_PER_CM < W) ticks.push(<span key={`q${i}${q}`} className="wd-ruler-tick" style={{ left: x + q * PX_PER_CM }} />);
  }

  const leftBound = pos('margin', 0, ML);
  const rightBound = pos('margin', 1, W - MR);
  const first = para ? para.first : ML;
  const left = para ? para.left : ML;
  const right = para ? para.right : W - MR;
  const edge = para ? para.edge : ML;
  const bar = drag && ref.current ? ref.current.getBoundingClientRect() : null;

  return (
    <div className="wd-ruler" ref={ref} style={{ width: W }} aria-hidden="true" data-tab-type={tabType}>
      <div className="wd-ruler-margin" style={{ left: 0, width: leftBound }} />
      <div className="wd-ruler-margin" style={{ left: rightBound, width: W - rightBound }} />
      <div className="wd-ruler-band" style={{ left: ML, width: Math.max(0, W - ML - MR) }} onMouseDown={addStop} />
      {ticks}
      <button
        type="button"
        className="wd-ruler-type"
        data-type={tabType}
        data-tip={`Tab stop to add: ${tabType}. Click to change.`}
        onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
        onClick={() => setTabType(TAB_TYPES[(TAB_TYPES.indexOf(tabType) + 1) % TAB_TYPES.length])}
      >
        <i className="wd-ruler-glyph" data-align={tabType} />
      </button>
      <div
        className="wd-ruler-bound"
        data-bound="left"
        data-tip="Left margin"
        style={{ left: leftBound - 3 }}
        onMouseDown={(e) => onMargin && hold(e, { kind: 'margin', index: 0, x: ML, min: 24, max: W - MR - 120, release: (x) => onMargin('left', x) })}
      />
      <div
        className="wd-ruler-bound"
        data-bound="right"
        data-tip="Right margin"
        style={{ left: rightBound - 3 }}
        onMouseDown={(e) => onMargin && hold(e, { kind: 'margin', index: 1, x: W - MR, min: ML + 120, max: W - 24, release: (x) => onMargin('right', W - x) })}
      />
      {para && onParagraph ? (
        <>
          <div
            className="wd-ruler-first"
            data-marker="first"
            data-tip="First line indent"
            style={{ left: pos('first', 0, first) - 5 }}
            onMouseDown={(e) => hold(e, {
              kind: 'first', index: 0, x: first, min: edge, max: right - 8,
              release: (x, dx) => {
                const ti = para.indentPx + dx;
                onParagraph(ti >= 0 ? { firstLineTwips: tw(ti) } : { hangingTwips: tw(-ti) });
              },
            })}
          />
          <div
            className="wd-ruler-hang"
            data-marker="hanging"
            data-tip="Hanging indent"
            style={{ left: pos('hang', 0, left) - 5 }}
            onMouseDown={(e) => hold(e, {
              kind: 'hang', index: 0, x: left, min: edge, max: right - 8,
              release: (x, dx) => {
                // The lines after the first move; the first line stays.
                const ti = para.indentPx - dx;
                onParagraph({ leftTwips: tw(Math.max(0, para.leftPx + dx)), ...(ti >= 0 ? { firstLineTwips: tw(ti) } : { hangingTwips: tw(-ti) }) });
              },
            })}
          />
          <div
            className="wd-ruler-left"
            data-marker="left"
            data-tip="Left indent"
            style={{ left: pos('left', 0, left) - 5 }}
            onMouseDown={(e) => hold(e, {
              kind: 'left', index: 0, x: left, min: edge + Math.max(0, -para.indentPx), max: right - 8,
              release: (x, dx) => onParagraph({
                leftTwips: tw(Math.max(0, para.leftPx + dx)),
                ...(para.indentPx > 0 ? { firstLineTwips: tw(para.indentPx) } : para.indentPx < 0 ? { hangingTwips: tw(-para.indentPx) } : {}),
              }),
            })}
          />
          <div
            className="wd-ruler-right"
            data-marker="right"
            data-tip="Right indent"
            style={{ left: pos('right', 0, right) - 5 }}
            onMouseDown={(e) => hold(e, {
              kind: 'right', index: 0, x: right, min: Math.max(first, left) + 8, max: W - MR,
              release: (x, dx) => onParagraph({ rightTwips: tw(Math.max(0, para.rightPx - dx)) }),
            })}
          />
          {stops.map((s, i) => {
            const x = ML + s.posPx;
            return (
              <div
                key={i}
                className={`wd-ruler-tab${drag && drag.kind === 'tab' && drag.index === i && drag.off ? ' off' : ''}`}
                data-align={s.align || 'left'}
                data-pos={Math.round(s.posPx)}
                data-tip={`${s.align || 'left'} tab at ${(s.posPx / PX_PER_CM).toFixed(2)} cm — drag to move, drag off to remove`}
                style={{ left: pos('tab', i, x) - 4 }}
                onMouseDown={(e) => hold(e, {
                  kind: 'tab', index: i, x, min: ML + 1, max: W - MR,
                  release: (nx, dx) => {
                    if (nx == null) writeStops(stops.filter((_, j) => j !== i));
                    else writeStops(stops.map((t, j) => (j === i ? { ...t, posPx: t.posPx + dx } : t)));
                  },
                })}
              />
            );
          })}
        </>
      ) : null}
      {table && onColumn ? table.xs.map((c, i) => {
        const prev = i === 0 ? table.left : table.xs[i - 1].x;
        const next = table.xs[i + 1]?.x ?? null;
        return (
          <div
            key={c.k}
            className="wd-ruler-col"
            data-k={c.k}
            data-tip="Table column — drag to resize"
            style={{ left: pos('col', c.k, c.x) - 4 }}
            onMouseDown={(e) => hold(e, {
              kind: 'col', index: c.k, x: c.x, min: prev + MIN_COLUMN, max: next != null ? next - MIN_COLUMN : W,
              release: (x, dx) => onColumn(tableId, c.k, dx / table.scale),
            })}
          />
        );
      }) : null}
      {drag && bar ? (
        <>
          <div className="wd-ruler-guide" style={{ left: bar.left + drag.x, top: bar.bottom }} />
          <div className="wd-ruler-readout" style={{ left: drag.x }}>
            {drag.off ? 'Remove' : `${((drag.x - ML) / PX_PER_CM).toFixed(2)} cm`}
          </div>
        </>
      ) : null}
    </div>
  );
}

/**
 * Grips on the table at the caret: an invisible strip on every column edge
 * and every row's bottom, the pointer changing to say so; a drag draws a
 * dashed line where the border will go and, on release, tells the engine
 * the two columns either side (a column's width, in grid px) or the row's
 * new height. `onDrag` lets the page know a drag is on, so the mouse-up
 * that ends it is not read as a caret move.
 */
export function TableGrips({ page, model, pages, tableId, gridPx, onColumn, onRow, onDrag }) {
  const [parts, setParts] = React.useState(null);
  const [guide, setGuide] = React.useState(null);
  const measure = React.useCallback(() => setParts(measureTable(page.current, tableId, gridPx)), [page, tableId, gridPx]);
  React.useLayoutEffect(() => {
    measure();
  }, [measure, model, pages]);
  useResize(page, measure);
  if (!parts) return null;

  const hold = (e, spec) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const start = spec.axis === 'x' ? e.clientX : e.clientY;
    let at = spec.at;
    onDrag?.(true);
    setGuide({ axis: spec.axis, at, part: spec.part });
    const move = (ev) => {
      at = clamp(spec.at + ((spec.axis === 'x' ? ev.clientX : ev.clientY) - start), spec.min, spec.max);
      setGuide({ axis: spec.axis, at, part: spec.part });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setGuide(null);
      setTimeout(() => onDrag?.(false), 0);
      const d = Math.round(at - spec.at);
      if (d !== 0) spec.release(d);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  return (
    <div className="wd-tgrips" contentEditable={false} aria-hidden="true">
      {parts.map((part, p) => (
        <React.Fragment key={p}>
          {part.xs.map((c, i) => {
            const prev = i === 0 ? part.left : part.xs[i - 1].x;
            const next = part.xs[i + 1]?.x ?? null;
            return (
              <div
                key={`c${c.k}`}
                className="wd-tgrip col"
                data-k={c.k}
                style={{ left: c.x - 3, top: part.top, height: part.height }}
                onMouseDown={(e) => hold(e, {
                  axis: 'x', at: c.x, part, min: prev + MIN_COLUMN, max: next != null ? next - MIN_COLUMN : c.x + 600,
                  release: (d) => onColumn(tableId, c.k, d / part.scale),
                })}
              />
            );
          })}
          {part.ys.map((r) => (
            <div
              key={`r${r.r}`}
              className="wd-tgrip row"
              data-r={r.r}
              style={{ top: r.y - 3, left: part.left, width: part.width }}
              onMouseDown={(e) => hold(e, {
                axis: 'y', at: r.y, part, min: r.top + MIN_ROW, max: r.y + 600,
                release: (d) => onRow(tableId, r.r, r.y - r.top + d),
              })}
            />
          ))}
        </React.Fragment>
      ))}
      {guide ? (
        <div
          className={`wd-tguide ${guide.axis}`}
          style={guide.axis === 'x'
            ? { left: guide.at, top: guide.part.top, height: guide.part.height }
            : { top: guide.at, left: guide.part.left, width: guide.part.width }}
        />
      ) : null}
    </div>
  );
}

const CSS = `
/* the ruler ---------------------------------------------------------------- */
.wd-ruler {
  position: relative; flex: none; height: 22px; margin: 10px auto 0; box-sizing: border-box;
  background: var(--surface); border: 1px solid var(--line); border-radius: 3px 3px 0 0;
  font-size: 9px; color: var(--ink-3); user-select: none;
}
.wd-ruler-margin { position: absolute; top: 0; bottom: 0; background: var(--sunken); }
.wd-ruler-band { position: absolute; top: 0; bottom: 0; cursor: default; }
.wd-ruler-tick { position: absolute; bottom: 4px; width: 1px; height: 3px; background: var(--ink-3); pointer-events: none; opacity: .8; }
.wd-ruler-tick.half { height: 5px; }
.wd-ruler-tick.major { width: 14px; margin-left: -7px; height: auto; bottom: 3px; background: none; text-align: center; line-height: 1; opacity: 1; }
.wd-ruler-first, .wd-ruler-hang, .wd-ruler-right {
  position: absolute; width: 0; height: 0; border-left: 5px solid transparent; border-right: 5px solid transparent;
  cursor: ew-resize; z-index: 2;
}
.wd-ruler-first { top: 1px; border-top: 6px solid var(--ink-2); }
.wd-ruler-hang, .wd-ruler-right { top: 9px; border-bottom: 6px solid var(--ink-2); }
.wd-ruler-left { position: absolute; top: 15px; width: 10px; height: 5px; background: var(--ink-2); cursor: ew-resize; z-index: 2; }
.wd-ruler-first:hover, .wd-ruler-hang:hover, .wd-ruler-right:hover { filter: brightness(1.3); }
.wd-ruler-left:hover, .wd-ruler-tab:hover, .wd-ruler-col:hover { filter: brightness(1.3); }
.wd-ruler-tab, .wd-ruler-glyph {
  position: absolute; top: 7px; width: 8px; height: 8px; box-sizing: border-box; cursor: ew-resize; z-index: 2;
  border-bottom: 2px solid var(--ink-2);
}
.wd-ruler-tab[data-align="left"], .wd-ruler-glyph[data-align="left"] { border-left: 2px solid var(--ink-2); }
.wd-ruler-tab[data-align="right"], .wd-ruler-glyph[data-align="right"] { border-right: 2px solid var(--ink-2); }
.wd-ruler-tab[data-align="center"]::after, .wd-ruler-tab[data-align="decimal"]::after,
.wd-ruler-glyph[data-align="center"]::after, .wd-ruler-glyph[data-align="decimal"]::after {
  content: ''; position: absolute; left: 3px; top: 0; width: 2px; height: 6px; background: var(--ink-2);
}
.wd-ruler-tab[data-align="decimal"]::before, .wd-ruler-glyph[data-align="decimal"]::before {
  content: ''; position: absolute; left: 6px; top: 3px; width: 2px; height: 2px; background: var(--ink-2);
}
.wd-ruler-tab.off { opacity: .35; }
.wd-ruler-col {
  position: absolute; top: 4px; width: 8px; height: 13px; box-sizing: border-box; cursor: ew-resize; z-index: 2;
  background: var(--surface-2); border: 1px solid var(--ink-3); border-radius: 1px;
  background-image: linear-gradient(var(--ink-3), var(--ink-3)); background-size: 1px 100%; background-position: center; background-repeat: no-repeat;
}
.wd-ruler-bound { position: absolute; top: 0; bottom: 0; width: 6px; cursor: ew-resize; z-index: 1; }
.wd-ruler-type {
  position: absolute; left: 3px; top: 3px; width: 15px; height: 15px; padding: 0; box-sizing: border-box;
  border: 1px solid var(--line); border-radius: 2px; background: var(--surface); cursor: pointer; z-index: 2;
}
.wd-ruler-type .wd-ruler-glyph { left: 3px; top: 3px; cursor: pointer; }
.wd-ruler-guide { position: fixed; width: 0; height: 100vh; border-left: 1px dashed var(--ink-3); pointer-events: none; z-index: 30; }
.wd-ruler-readout {
  position: absolute; top: -21px; transform: translateX(-50%); background: var(--ink); color: var(--surface);
  font-size: 10px; padding: 1px 5px; border-radius: 3px; white-space: nowrap; pointer-events: none; z-index: 3;
}

/* the grips on a table's borders ------------------------------------------- */
.wd-tgrips { position: absolute; inset: 0; pointer-events: none; z-index: 3; }
.wd-tgrip { position: absolute; pointer-events: auto; }
.wd-tgrip.col { width: 7px; cursor: col-resize; }
.wd-tgrip.row { height: 7px; cursor: row-resize; }
.wd-tguide { position: absolute; pointer-events: none; }
.wd-tguide.x { width: 0; border-left: 1px dashed var(--accent); }
.wd-tguide.y { height: 0; border-top: 1px dashed var(--accent); }
`;

let installed = false;
export function installRulerStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-ruler-css';
  style.textContent = CSS;
  document.head.appendChild(style);
}
