/**
 * Chart data as TEXT — one `label, value` per line.
 *
 * A GRID OF INPUTS IS WHAT A SPREADSHEET DOES, and it is the wrong instrument
 * here: this is a media tool, the charts are four to eight points, and typing
 * "Mon, 12" is faster than tabbing between two cells. It also PASTES — a column
 * copied out of a spreadsheet lands correctly with no import step, which is how
 * the numbers actually arrive.
 *
 * Round-tripping is lossless for the pairs it holds, so the textarea is rebuilt
 * from the stored arrays and the document never has to keep a second copy of
 * the same numbers.
 *
 * It lives HERE rather than beside the panel that uses it because it is pure
 * data logic with no notion of a control — and because a `.js` file full of JSX
 * cannot be imported by a Node test, which is exactly what this needs.
 */

/** The stored arrays → the text a person edits. */
export function dataToText(patch) {
    const labels = Array.isArray(patch.labels) ? patch.labels : [];
    // Reads the series shape too, so switching a bound chart back to
    // hand-typed data shows the numbers it already had.
    const values = Array.isArray(patch.values) ? patch.values
        : (Array.isArray(patch.series) && patch.series[0]?.values) || [];
    return values.map((v, i) => (labels[i] ? `${labels[i]}, ${v}` : String(v))).join("\n");
}

/**
 * The trailing number on a line, allowing thousands separators.
 *
 * A FIRST VERSION SPLIT ON THE LAST COMMA and it was wrong in a way only a test
 * found: "A, £1,200" splits there into the label "A, £1" and the value 200. Two
 * wanted behaviours collide — labels may contain commas ("Jan, Feb"), and a
 * pasted number may too — and no split point satisfies both. Matching the
 * NUMBER at the end and treating everything before it as the label satisfies
 * both, because the number is the part with a shape.
 *
 * A comma inside the number is read as a thousands separator, which makes
 * "3,5" thirty-five rather than three-and-a-half. That is a real ambiguity with
 * no context to resolve it; the documented format is "label, value" in English,
 * so the separator reading is the one that matches what the field says it takes.
 */
const TRAILING_NUMBER = /([-+]?\d[\d,]*(?:\.\d+)?(?:[eE][-+]?\d+)?)\s*%?\s*$/;

/** The text a person edited → the stored arrays. */
export function textToData(text) {
    const labels = [];
    const values = [];
    for (const line of String(text || "").split("\n")) {
        if (!line.trim()) continue;

        const m = TRAILING_NUMBER.exec(line);
        let label;
        let n;
        if (m) {
            n = Number(m[1].replace(/,/g, ""));
            // Trim what separated the label from the number, plus a currency
            // mark left clinging to it — "A, £" is not a label anybody typed.
            label = line.slice(0, m.index).replace(/[\s,:;$£€¥₹-]+$/, "").trim();
        } else {
            // No number on the line at all. Keep whatever label there is and
            // record a zero, rather than discarding the row: a typo in one
            // value should not silently delete its category.
            const at = line.lastIndexOf(",");
            label = (at >= 0 ? line.slice(0, at) : "").trim();
            n = 0;
        }

        labels.push(label);
        values.push(Number.isFinite(n) ? n : 0);
    }
    return { labels, values };
}
