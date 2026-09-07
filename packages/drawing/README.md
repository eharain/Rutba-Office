# @rutba/drawing

Charts, shapes and drawings as data. A JSON scene graph, an SVG renderer, a chart
layout engine, and a colour palette that was validated rather than chosen.

Built for three consumers from the start — Workspace (charts in a sheet), Studio
(the design surface) and Mail (a chart in an email) — because retrofitting
shareability is how you end up with three chart libraries that disagree.

Nothing here imports a product, a DOM, a filesystem or a file format. A scene is
JSON and a render is a string, so the same code runs in a browser, on a server
and in a test.

```js
import { buildChart, renderSvg } from '@rutba/drawing';

const svg = renderSvg(buildChart({
  type: 'column',
  title: 'Revenue by region',
  categories: ['North', 'South', 'East'],
  series: [{ name: '2026', values: [1420, 860, 240] }],
  width: 480, height: 300,
  mode: 'dark',            // a selected palette, not an inverted one
}));
```

## Layers, weakest coupling first

| Module | What it is |
|---|---|
| `palette.js` | the validated colour system — the one thing all consumers must share |
| `scene.js` | the JSON scene graph |
| `measure.js` | text metrics without a DOM |
| `svg.js` | scene → SVG string |
| `chart.js` | chart spec → scene, with the design rules enforced |
| `shapes.js` | shape/picture descriptor → scene |
| `ooxml.js` | DrawingML adapter, at the edge and optional |

## Rules the API enforces, so you cannot get them wrong

- **No dual axis.** Two y-scales is the most common way to mislead with a chart,
  and it is simply not expressible.
- **Colour follows the entity, never its rank.** Filtering a series out does not
  repaint the survivors.
- **Hues are assigned in fixed order and never cycled.** A ninth series returns
  muted ink rather than an invented hue.
- **Scatter is capped at three series** — beyond that the all-pairs colour
  separation fails. It throws, with the reason.
- **A legend appears for two or more series and never for one.**
- **Text wears ink tokens, never the series colour.**

## The palette is not decoration

Colour here is an accessibility property with a pass/fail answer, and the values
were validated, not picked. The measurements are recorded in `palette.js` so a
future change can be compared rather than guessed. If you change a hex, re-run
the validator. The slot **order** is the colour-blind-safety mechanism.

## For email, honestly

Several major clients — Outlook on Windows especially — do not render SVG at all.
For email this renderer produces the *source*, not the delivery format: an
emailed chart needs a raster step to PNG that does not exist in this process.
`RASTER_REQUIRED_NOTE` is exported so nobody ships one assuming otherwise.
**That gap is real and unsolved.**

## Status

Plain ESM JavaScript, no build step, no runtime dependencies. Lives in
Rutba-Workspace for now; belongs in Rutba-Platform once a second consumer
arrives. See `BOUNDARIES.md` and `SHARED.md` in Rutba-Workspace.
