# Text recognition

Reading the words in a picture or a scanned page, on this computer, with
the network off. This page records how well the engine reads each of the
suite's six languages, measured before any of it was built into the apps,
and how to measure it again.

## The engine

Tesseract 5 compiled to WebAssembly, as tesseract.js 7 ships it
(tesseract.js-core 7.0.0, Apache-2.0), with the "fast" LSTM models of
tessdata_fast (Apache-2.0) for English (`eng`), Urdu (`urd`), Arabic
(`ara`), Hindi (`hin`), Chinese, Simplified (`chi_sim`) and Japanese
(`jpn`). Both are vendored in the repository and ship inside the suite, so
nothing is fetched at first use or ever. The SIMD, LSTM-only build of the
engine is used: 2.9 MB of WebAssembly. The six models are 13.0 MB:

| Model | Size |
| :-- | --: |
| eng | 4.1 MB |
| chi_sim | 2.5 MB |
| jpn | 2.5 MB |
| ara | 1.4 MB |
| urd | 1.4 MB |
| hin | 1.1 MB |

## Where it runs

`@rutba/ocr` (packages/ocr) answers `recognize(image, { languages })` with
the text, its lines and its words, each word with its box in the picture's
pixels and a confidence from 0 to 100. Pages are read on a pool of worker
threads (`src/host.js`): one page a thread, at most one thread a core and
never more than four, the rest waiting in a queue in the order asked, so a
long scan holds nothing else up; a job can be cancelled while it waits or
while it is read. Nothing is read on the main process or on a window's
thread. The windows reach it through the `ocr` namespace
(apps/desktop/main/ocr.js): `recognize`, `cancel`, `languages` and
`stats`. A window that closes cancels what it was waiting for.

The engine and the models are vendored in `packages/ocr/vendor` by
`tools/vendor-ocr.mjs` and unpacked beside the installed app's archive,
where the threads read them. The engine is handed its WebAssembly as bytes
and its models are written into its own memory, so it has nowhere to fetch
anything from; the package's tests run a recognition with every way to the
network refused, in the process and its threads, and check that none was
tried.

## How it was measured

**Corpus.** `tools/ocr-corpus.js` sets four sentences of each language
(`tools/ocr-texts.json`: a notice, an invoice line with numbers, a request
and a description) as one page of 12 point type in a 6.5 inch column, drawn
by the suite's own renderer (Chromium in Electron), so the text of every
page is known exactly. Each page is made three ways: clean at 150 dots per
inch, clean at 300, and at 300 as a simulated scan (turned 1.2 degrees,
softened by half a pixel, on a paper grain, saved as a JPEG at quality 55).
Fonts: Times New Roman and Arial for English, Noto Naskh Arabic for Arabic
and for Urdu in Naskh, Noto Nastaliq Urdu for Urdu in Nastaliq, Noto Sans
Devanagari for Hindi, SimSun for Chinese and Yu Gothic for Japanese. 24
pages in all.

**Rates.** The character error rate (CER) is the edit distance between the
text read and the text drawn, over the length of the text drawn, counted in
Unicode code points after NFC and with runs of white space made one space.
The word error rate (WER) is the same over words. Chinese and Japanese are
written without spaces, so their CER ignores spaces and they have no WER.

**Time and memory.** `tools/ocr-measure.mjs` reads each page in a Node
process of its own, twice. "First" is the time from nothing to the first
page's text: starting the engine, loading the model and reading the page.
"Page" is the second read alone, the cost of each further page. Peak
memory is the process's peak resident set over both. The machine is the
reference laptop: an Intel Core i5-8350U (four cores, eight threads, 1.7
GHz), 16 GB of memory, Windows 11, Node 22. Other work ran on it at the same
time, so single times vary by a third or more; the medians are steadier.

## Results

Fast models, SIMD build. CER and WER in percent; time per further page.

| Page | Clean, 150 dpi | Clean, 300 dpi | Simulated scan, 300 dpi |
| :-- | :-- | :-- | :-- |
| English, serif | 0.2 / 2.2, 1.1 s | 0.0 / 0.0, 1.2 s | 0.0 / 0.0, 1.3 s |
| English, sans | 0.2 / 2.2, 0.8 s | 0.2 / 2.2, 1.2 s | 0.0 / 0.0, 1.5 s |
| Arabic, Naskh | 1.0 / 5.7, 0.8 s | 3.2 / 8.6, 1.0 s | 3.4 / 10.0, 1.1 s |
| Urdu, Naskh | 4.0 / 7.5, 1.1 s | 5.5 / 8.6, 1.1 s | 7.3 / 10.8, 1.1 s |
| Urdu, Nastaliq | 20.9 / 64.5, 0.7 s | 19.4 / 62.4, 0.9 s | 50.0 / 86.0, 1.5 s |
| Hindi | 0.0 / 0.0, 1.2 s | 0.0 / 0.0, 1.8 s | 0.0 / 0.0, 1.9 s |
| Chinese | 1.5, 1.3 s | 2.9, 1.4 s | 2.2, 1.5 s |
| Japanese | 0.0, 1.1 s | 0.0, 1.8 s | 3.1, 1.5 s |

- **Time.** A page of this size takes about a second once the engine is
  running (median 1.2 s) and about two seconds from a cold start (median
  1.8 s, of which starting the engine and loading a model is 0.2 to 0.6 s).
  The engine reads on one thread; the pool reads pages on up to four
  threads at once, so a long scan goes up to four times as fast.
- **Memory.** 100 to 135 MB of peak resident memory per process for one
  page, Node's own share included; the engine and one model hold roughly
  40 to 80 MB of it. Four engines reading at once would stay under half a
  gigabyte.
- **English, Hindi and Japanese** are read almost without error on clean
  pages and simulated scans alike.
- **Arabic and Urdu in Naskh** are read well. Most of the errors are
  punctuation: the Arabic comma read as a Latin comma, or lost. Without
  punctuation the rates are 0.0 to 2.3 percent (CER) for Arabic and 2.9 to
  6.1 percent for Urdu.
- **Urdu in Nastaliq**, the style most Urdu is printed in, is read poorly:
  one character in five is wrong on a clean page, and three words in five;
  half the characters on a simulated scan. Urdu recognition is therefore
  labelled beta in the suite, good on Naskh and on Urdu typed in a Naskh
  font, not yet for Nastaliq print. A recogniser of our own for Nastaliq is
  planned.
- **Chinese** loses a character or two a page (1.5 to 2.9 percent).
- **150 or 300 dots per inch.** The models read type of 20 to 30 pixels best;
  at 300 dots per inch 12 point type is 50 pixels high and some scripts read
  a little worse than at 150, so scaling a page's text towards that size
  before reading is worth doing.

### The engine's builds

Median time per further page over the 24 pages, fast models:

| Build | WebAssembly | Page | First page | CER |
| :-- | --: | --: | --: | :-- |
| SIMD, LSTM only (shipped) | 2.9 MB | 1.18 s | 1.82 s | as above |
| relaxed SIMD, LSTM only | 2.9 MB | 1.17 s | 1.67 s | identical |
| no SIMD, LSTM only | 2.9 MB | 2.99 s | 3.30 s | identical |

SIMD makes reading two and a half times faster. Relaxed SIMD adds nothing
measurable on this processor, and the plain SIMD build runs on more
machines, so it is the one shipped.

### Fast and "best_int" models

The integer-quantised "best" models (`4.0.0_best_int` from tesseract.js's
model packages) were measured on the same pages with the same build:

| | eng | ara | hin | jpn | chi_sim, urd | Six together |
| :-- | :-- | :-- | :-- | :-- | :-- | --: |
| Size, fast | 4.1 MB | 1.4 MB | 1.1 MB | 2.5 MB | the same file | 13.0 MB |
| Size, best_int | 5.2 MB | 2.5 MB | 1.7 MB | 3.0 MB | the same file | 16.3 MB |
| CER, fast | 0.0 to 0.2 | 1.0 to 3.4 | 0.0 | 0.0 to 3.1 | | |
| CER, best_int | 0.0 | 1.0 to 1.2 | 0.0 to 0.5 | 0.5 to 4.6 | | |

The best_int models read English and Arabic slightly better, Japanese
slightly worse, at the same speed and 3.3 MB more. The fast models are
shipped; English and Arabic could move to best_int later if real scans
show the same gain.

### Not measured yet

- Real scans and photographs of pages (receipts, letters, a newspaper
  column). The simulated scans stand in for them until a set of real pages
  with their text is gathered.
- A second engine for Chinese, Japanese and difficult layouts.

## Measuring again

```bash
npm install --prefix <scratch> tesseract.js@7
npx electron tools/ocr-corpus.js --fonts <folder with the Noto fonts> --out <corpus>
node tools/ocr-measure.mjs --corpus <corpus> --core <scratch>/node_modules/tesseract.js/node_modules/tesseract.js-core --models <tessdata_fast folder> --out results.json
```

The Noto fonts (SIL Open Font Licence) are not kept in the repository:
`NotoNaskhArabic.ttf`, `NotoNastaliqUrdu.ttf` and `NotoSansDevanagari.ttf`,
from the Noto project or Google Fonts. `--build` picks the engine's build
(`simd-lstm`, `relaxedsimd-lstm`, `lstm`) and `--only` a page name's start.
