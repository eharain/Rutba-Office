# Releasing

How a version of Rutba Office is numbered, proven and published. The order
here is the order it is done in.

## The number

- **A feature takes the minor number.** A release that adds anything a
  person can do that they could not before is the next minor: 1.29.10 was
  followed by 1.30.0, not 1.29.11.
- **A patch carries fixes only.** 1.30.1 puts right what 1.30.0 got wrong,
  and adds nothing else. A fix that needs a new control waits for the next
  minor.
- The updater and a person reading the number can then tell the two apart,
  which ROADMAP.md found they could not across 1.29.1 to 1.29.10.

The version is in three places, moved together: `package.json`,
`apps/desktop/package.json` and the lock file's two entries for them
(`npm install --package-lock-only` brings the lock file along).

## Before the number is moved

1. Each piece of work is on a branch of its own, off an up-to-date `main`,
   one theme to a branch, its history linear. Fetch first: `main` moves
   between releases.
2. `npm test` passes on the branch.
3. The window checks for what changed pass on the branch:
   `RUTBA_VERIFY_ONLY=<keys> npm run verify:apps`.
4. The branches are put together, in order, on one integration branch.

## The note

`docs/releases/v<version>.md`: a first line that says what the release is,
then a section per app, then **Fixed** and anything else the release needs.
Wrapped at about 75 characters; `tools/release-notes.js` joins the lines
back together for the release page. The note is plain: no en or em dashes.
`docs/GAPS.md` and the README are brought up to date in the same commit.

## The gate

```
npm run gate -- --record
```

Run on Windows, on the release commit, with nothing else running against
the same profile. It runs the engine suite, the editing checks, every
application check and the smoke captures. With `--record`, a full run that
passes writes its own counts into the note's **Gate** section (where it
ran, when, and what each pass counted) and into the README's counts, so the
release says the gate ran rather than leaving it to be believed. Commit
those two files with the version bump.

A check that fails because of the machine (a stand-in camera, a font the
system lacks) is run again on its own; if it passes alone it is named in
the release commit's message. A check that fails twice is a fault, and the
release waits for it.

The fuzzer is run over the readers before a minor release, and its count
goes in the note:

```
node tools/fuzz-open.js 300 <seed>
```

It damages a new document, workbook and deck and every old-format file in
`tests/fixtures/binary`, opens each in a worker of its own with a
two-second limit, and fails on anything that is not a document, a password
asked for, or a refusal in a sentence.

## Publishing

1. `git fetch`; the integration branch must have `origin/main` in it.
2. `main` and `dev` are moved to the release commit (`git fetch . <branch>:main
   <branch>:dev`); nothing is merged on the way.
3. `git tag -a v<version> -m "Rutba Office <version> beta"`.
4. `git push origin main dev v<version>`.
5. The tag starts `.github/workflows/release.yml`, which runs the engine
   suite on Windows, builds the installer and the portable executable, and
   publishes them with the note (`tools/publish-release.js`). Watch it to
   the end (`gh run watch`) and look at the release's files
   (`gh release view v<version>`): the installer, the portable executable,
   `latest.yml` and the blockmap.
6. The site (Rutba-Management) is brought to the release in the same sitting:
   `npm run import:releases`, then `npm run check:facts`, there.
7. The branches that went into the release are deleted locally.
