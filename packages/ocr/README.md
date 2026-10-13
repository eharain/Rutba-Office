# @rutba/ocr

Text recognition on this computer.

```js
import { recognize } from '@rutba/ocr';
const { text, lines, words } = await recognize(pixels, { languages: ['urd', 'eng'] });
// words: [{ text, box: { x, y, width, height }, confidence, line }]
```

`pixels` is `{ width, height, data }` with one to four eight-bit channels (a
canvas's ImageData will do) or a PNG file's bytes. Languages: `eng`, `urd`
(beta), `ara`, `hin`, `chi_sim`, `jpn`.

Recognition runs on a pool of worker threads (`src/host.js`), one page a
thread, at most one thread a core and never more than four, the rest
queued; a job can be cancelled with an `AbortSignal`. It never runs on the
caller's thread. The engine and the models are read from `vendor/`, inside
the package, and nothing is fetched: `test/ocr.test.js` runs a recognition
with every way to the network refused and checks none was tried.

docs/OCR.md has the measured error rates, speed and memory.

## Notice

`vendor/core` is Tesseract 5 compiled to WebAssembly as tesseract.js 7 ships
it (tesseract.js-core 7.0.0, Apache-2.0); `vendor/tessdata` holds the fast
LSTM models of tesseract-ocr/tessdata_fast (Apache-2.0). Each folder carries
its licence and a VERSION file with every file's SHA-256.
`tools/vendor-ocr.mjs` writes both.
