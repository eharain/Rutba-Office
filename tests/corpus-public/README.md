# The starter corpus

A few files of every format the corpus run reads (`npm run verify:corpus`,
see `tools/corpus.js` and `docs/COMPATIBILITY.md`): Word, Excel and
PowerPoint in their XML and 97-2003 forms, OpenDocument, PDF, mail
messages, an Outlook data file, a calendar and contact files.

Every file here was made by this project: written by the suite's own
writers from words written here, or copied from the repository's test
fixtures, some of which were saved by Microsoft Office from content written
here. Nothing was downloaded and nothing belongs to anyone else, so every
file is under the repository's licence, AGPL-3.0-or-later.
`SOURCES.json` names each file's source, licence and SHA-256, and
`tests/corpus.test.js` fails when a file is here that it does not name.

To make the folder again: `node tools/make-corpus.mjs`.

The real-world corpus the release criteria ask for is people's own files.
It is gathered outside the repository (`.corpus/` is ignored) and run with
`RUTBA_CORPUS=<folder> npm run verify:corpus`.
