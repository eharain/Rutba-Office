# Compatibility: the corpus run

How Rutba Office is measured against real files, and the shape of the
results a compatibility table is made from.

## The run

```
npm run verify:corpus
```

`tools/corpus.js` opens every file under a folder in the engines, the same
document service, PDF reader, mail, calendar and contact readers the
application runs, each file in a worker thread of its own with a deadline.
For each file it records whether it opened, whether what was read survived
a save and a second open, what the engines warned about, any fault, and how
long each step took.

| Setting | What it does |
|---|---|
| `RUTBA_CORPUS=<folder>` | The corpus. Every file of a known format in any sub-folder. Default: `tests/corpus-public`, the starter corpus. |
| `RUTBA_CORPUS_OUT=<folder>` | Where `results.json` and `summary.md` go. Default: `apps/desktop/build/corpus` (ignored by git). |
| `RUTBA_CORPUS_WINDOWS=1` | Also open each file in the application's windows (`apps/desktop/main/verify-corpus.js`), one window at a time, and add what each window did to its result. Needs Electron, so run it through the shared lock. |
| `RUTBA_CORPUS_RATES=<file.json>` | The agreed rates: `{ "docx": { "open": 0.98, "roundTrip": 0.95 }, "*": { "open": 0.9 } }`. A format below its rate fails the run. |
| `RUTBA_CORPUS_TIMEOUT=<ms>` | How long one file may take before it counts as a hang. Default 60000. |
| `RUTBA_CORPUS_LIMIT=<n>` | The first n files only. |

The run fails (exit code 1) when any file faults or hangs, or a format falls
below its agreed rate. A refusal in a sentence is not a fault: it counts
against its format's open rate. A file protected by a password is counted
apart: asking for the password is the right answer, and a corpus carries no
passwords.

The formats, and what a round trip writes:

| Format | Read by | Round trip writes |
|---|---|---|
| docx, xlsx, pptx | the document service | the same format, by Save |
| odt, ods, odp | the document service | the same format, by Save As |
| doc, xls, ppt | the document service | docx, xlsx, pptx (what Save does with a 97-2003 file) |
| pdf | `@rutba/pdf` reader | a PDF of every page, as the PDF app's Save assembles one |
| eml | `@rutba/mailbox` | the message written out again (`rfc822.js`) |
| pst | `@rutba/mailbox` (every folder, every message's headers, the first 200 read whole) | none: the suite reads Outlook data files and does not write them |
| ics | `@rutba/calendar` | the calendar written out again |
| vcf | `@rutba/contacts` | the cards written out again |

The Data Files format (`.rbase`) is not in the corpus: it is the suite's
own, specified in [DATA-FILES-FORMAT.md](DATA-FILES-FORMAT.md), and
`tools/check-rbase.mjs` says whether a file follows it.

"Round-tripped equal" compares the suite's reading of the file before and
after: for documents the text of every paragraph of the body in order and
the number of tables and drawings; for workbooks the sheet names in order
and every cell's value, or its formula where it has one; for decks each
slide's text and pictures; for PDF each page's text; for mail the subject,
sender, recipients, text and each attachment's name and size; for calendars
each event's UID, summary, start, end, all-day flag and rule; for contacts
each card's name, organisation, e-mail and telephone. For a docx, xlsx or
pptx the "before" is the file itself as it is on disk; for the other office
formats it is the suite's first reading saved as the XML format. A reading
that is wrong the same way twice compares equal: this measures whether
saving loses what was read. How faithfully a file is drawn is measured by
the window checks and the fidelity tests.

## The starter corpus

`tests/corpus-public` holds a few files of every format, all made by this
project (`node tools/make-corpus.mjs`), with each file's source, licence and
SHA-256 in `SOURCES.json`. `tests/corpus.test.js` holds the folder to that
record. The real-world corpus, people's own files, never enters the
repository: it is gathered under `.corpus/` (ignored) or anywhere else, and
run with `RUTBA_CORPUS`.

## results.json

Schema `rutba-office-corpus/1`. A change of meaning to any field moves the
number.

```json
{
  "schema": "rutba-office-corpus/1",
  "suite": "2.0.0",
  "commit": "0e5225c",
  "ranAt": "2026-10-15T11:19:52.543Z",
  "platform": "win32-x64",
  "node": "v22.20.0",
  "corpus": { "root": "tests/corpus-public", "files": 40 },
  "windows": false,
  "formats": {
    "docx": {
      "files": 7,
      "opened": 6,
      "locked": 1,
      "refused": 0,
      "faults": 0,
      "timeouts": 0,
      "roundTrips": 6,
      "roundTripEqual": 6,
      "withWarnings": 0,
      "windowShown": 0,
      "windowNotShown": 0,
      "openRate": 1,
      "roundTripRate": 1,
      "ms": { "median": 13, "p90": 60, "max": 159 }
    }
  },
  "files": [
    {
      "file": "docx/letter.docx",
      "format": "docx",
      "bytes": 2051,
      "sha256": "9f2c...",
      "status": "opened",
      "opened": true,
      "converted": null,
      "roundTrip": { "target": "docx", "equal": true, "differences": [] },
      "warnings": [],
      "faults": [],
      "refused": null,
      "facts": { "paragraphs": 5, "tables": 0, "drawings": 0 },
      "ms": { "open": 9, "roundTrip": 12, "total": 21 },
      "window": null
    }
  ],
  "rateFailures": []
}
```

### Per file

| Field | Meaning |
|---|---|
| `file` | The path under the corpus folder, with forward slashes. |
| `format` | The extension, lower case: one of the fourteen above. |
| `bytes`, `sha256` | The file's size and hash, so a result can be tied to the file without naming whose it is. |
| `status` | `opened`; `locked` (a password was asked for); `refused` (the engine said in a sentence why it would not open it); `fault` (an error the engine threw rather than chose: a TypeError, a RangeError, a stack overflow); `timeout` (past the deadline; the worker is stopped and the run goes on). |
| `opened` | `status === "opened"`. |
| `converted` | For a 97-2003 or OpenDocument file, the format it was read as and converted from (`word97`, `biff8`, `biff5`, `ppt`, `odt`, `ods`, `odp`); else null. |
| `roundTrip` | null where no round trip was tried (a format the suite does not write, a file that did not open). Else `target` (the format written), `equal` (true, or false), `differences` (up to ten sentences, the first places the two readings part, in reading order) and, where the save itself failed, `error`. |
| `warnings` | What the engines said on the console while working on the file, and notes such as "protected with a password: the pages were counted", "no events found". |
| `faults` | The faults, as `Name: message`. A round trip that faulted is here too, prefixed `round trip:`. |
| `refused` | The sentence, when `status` is `refused`. |
| `facts` | How much was read: paragraphs, tables and drawings; sheets and cells; slides, paragraphs and pictures; pages and characters; recipients, characters and attachments; events; cards; for an Outlook file its kind, folders, messages and how many were read whole. A round trip of nothing proves nothing, and this shows it. |
| `ms` | `open` (until the file was read and drawn as a model), `roundTrip` (save, open again, compare) and `total`, in milliseconds. |
| `window` | With `RUTBA_CORPUS_WINDOWS=1`: `{ status, ms, reason }` from the window run, `status` one of `ok`, `locked`, `refused`, `timeout`, `hung`, `crashed`, `skipped` (no window draws that format in the corpus check yet: mail, calendar, contacts, PDF). Else null. |

### Per format

`formats.<format>` counts the files by status, `faults` (files with at least
one), `roundTrips` (tried) and `roundTripEqual`, `withWarnings`, and with a
window run `windowShown` and `windowNotShown` (refused, timed out, hung or
crashed in a window). `openRate` is `opened / (files - locked)` and
`roundTripRate` is `roundTripEqual / roundTrips`, each to four places, or
null when there was nothing to divide. `ms` is the median, 90th percentile
and maximum of the total time of the files that opened.

### summary.md

The same per-format numbers as a Markdown table, with the files to look at
(every one that did not open, faulted, or came back different) under it.
This is the table the site can show: the format, how many files, the open
rate and the round-trip rate, and the corpus it was measured on.
