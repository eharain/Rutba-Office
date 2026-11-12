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
the same profile. It runs in lanes, so that what can run side by side does:

1. **build**: the windows bundled once, for every check process after it.
2. **test:perf**: the performance budgets, alone, while the machine is
   still (timed beside other work they measure the machine).
3. **side by side**: the engine suite, the editing checks and a window
   check process a chunk, the longest chunk first: as many at once as the
   processors allow (one to each two and a half, four at most), another
   started only while 1.5 GB of memory is free for it, and added as memory
   comes free; the gate says when it runs fewer and why (`--jobs` to say
   instead). Close browsers and other large programs before a release run:
   with 2.7 GB free, the 2.11.0 run had one lane and took over an hour. A chunk is an app, or
   a part of what every app shares: `shell`, `ribbon`, `look` (galleries,
   panels, icons), `theme` (polish, dark theme), `access` (accessibility, keyboard, screen
   reader, high contrast) and `languages` (Urdu, Arabic, Hindi, spelling).
   Each process has a clipboard and global shortcuts of its own
   (`machine-fakes.js`), so no process touches another's or the person's,
   and each logs to `.gate/logs/<chunk>.log`.
4. **machine**: alone, the parts that need the computer to themselves,
   marked `{ machine: true }` in verify-apps.js: the computer's own
   clipboard and global shortcuts checked through the system adapters
   (`verify-machine.js`), timing budgets, drags, and copies a page makes
   itself (Mail's message frame).
5. **again**: alone, after the rest, the parts whose checks failed above,
   in a process started just for them, a machine-lane failure in the
   machine lane (D-2026-10-21-13, the owner: "Queue the failed ones to run
   at the end"). Each part is numbered `<app>#<n>` in verify-apps.js
   (`<app>#main` for a section outside the parts), and
   `RUTBA_VERIFY_PART_IDS` runs only those named. The gate passes only if
   every check that failed is seen passing here by its name and nothing
   here fails; one that fails again closes the gate. The note's Gate section
   names each check that ran again.
6. **smoke**: every window paints without logging anything.

With `--record`, a full run that passes writes its own counts into the
note's **Gate** section (where it ran, when, and what each pass counted) and
into the README's counts, so the release says the gate ran rather than
leaving it to be believed. Commit those two files with the version bump.
`--serial` runs the window checks as one process in turn, as before.

A check that fails twice, beside the others and again alone, is a fault:
the gate is closed, and under the owner's rule the release is skipped and
its work goes into the next one.

The gate keeps a record of its last run in `.gate/` (each chunk's counts
and how long it took, which orders the next run):

```
npm run gate -- --resume              carry on: what passed on this commit is not run again
npm run gate -- --apps mail,pdf       the window checks of those chunks only
npm run gate -- --changed             the chunks changed since the last release tag
RUTBA_VERIFY_LANE=machine npm run verify:apps    the machine lane alone
```

After a failure, fix it and run `--resume`: only the passes and the chunks
that did not pass run again, and `--record` adds the runs up. The record
holds only while the commit and the working tree stay the same; a commit
starts it afresh. `--apps` and `--changed` are for working days, and a run
limited to some chunks is never put on the record.

The gate runs the fuzzers and the corpus run, and their counts go in the
note. To repeat a fuzzer's run, with the seed its log prints:

```
node tools/fuzz-open.js 300 <seed>
```

It damages a new document, workbook and deck and every old-format file in
`tests/fixtures/binary`, opens each in a worker with a two-second limit (a
new worker after any fault or time-out, so one bad file cannot take the next
with it), and fails on anything that is not a document, a password asked for,
or a refusal in a sentence.

The PDF reader has its own, since 1.46.0, because the PDF app reads files it
did not write on the main process:

```
node tools/fuzz-pdf.js 300 <seed>
```

## Publishing

Before the tag, run `.github/workflows/candidate.yml` from the Actions tab at
the commit the gate passed. It builds that commit and installs it over the last
release in each way a person's copy can be updated, the all-users copy among
them. Beside the build it runs the engine tests as the release does, on a clean
runner whose account is an administrator: a fault this PC cannot show (2.25.0's
Locker refused every open there) fails the candidate run, not the release; the
install jobs still run beside it, so read the run's own result. Installed
copies look for updates soon after they start, so this is the last time a fault
can be found before people have it.

1. `git fetch`; the integration branch must have `origin/main` in it.
2. `main` and `dev` are moved to the release commit (`git fetch . <branch>:main
   <branch>:dev`); nothing is merged on the way.
3. `git tag -a v<version> -m "Rutba Office <version>"`.
4. `git push origin main dev v<version>`.
5. The tag starts `.github/workflows/release.yml`, which runs the engine
   suite on Windows, builds the installer and the portable executable, and
   leaves them with the note in a draft (`tools/publish-release.js`). Watch
   it to the end (`gh run watch`) and look at the draft's files
   (`gh release view v<version>`): the installer, the portable executable,
   `latest.yml` and the blockmap.

   The release is made as a draft, which nobody sees and the updater's
   feed leaves out. Every file is uploaded to it, then `latest.yml` is read
   back from the release and held to the files (its version the tag's,
   each file it names uploaded at the size it says with the SHA-512 of the
   file built, the installer's blockmap beside it). A run that fails leaves
   the draft: run the workflow again and it finds the draft, uploads what
   is missing and checks again.

   **Signing and publishing, on the release PC only.** With
   `RUTBA_UPDATE_KEY_FILE` pointing at the update key:

   ```
   node tools/publish-release.js v<version> --sign
   ```

   It reads each installer back from the draft, holds it to `latest.yml`,
   signs it, uploads the `.sig` beside it and publishes. An installed copy
   installs nothing without a good `.sig`
   (`apps/desktop/main/update-signature.js`). The key is never put in
   Actions or in any repository. There are two keys, and each copy accepts
   either: losing both means no copy can update again, so both are backed
   up away from this PC. A release can be signed again if a `.sig` is lost.

   Steps 3 and 4 can leave the tag out: push `main` and `dev`, then
   `gh workflow run release.yml --ref main`. The release then makes the
   tag itself, at the commit it built, when `--sign` publishes it, so the
   tag never exists without its files. A pushed tag still starts the
   workflow as before; the release stays a draft until it is signed either
   way. When `--sign` has published it, `.github/workflows/upgrade.yml` runs by itself
   (see **The upgrade test**); its four jobs must pass.
6. The site (Rutba-Management) is brought to the release in the same sitting:
   `npm run import:releases`, then `npm run check:facts`, there.
7. The branches that went into the release are deleted locally.

## The upgrade test

Every release is proven to take over the previous one's profile:
`tools/upgrade-test.js` starts the older version on a profile seeded as a
person's would be (settings changed, a recent document, a setting no
version knows), waits until it has recorded itself there, closes it, and
starts the newer one on the same profile with
`RUTBA_OFFICE_UPGRADE_REPORT`, which writes what it read and quits
(`apps/desktop/main/upgrade-report.js`). The settings, the recent list and
the arrival record (the card that says which version replaced which) must
have survived, and the newer copy's updater must have read the release
feed. The judgement is `tools/upgrade-check.js`, tested in
`tests/upgrade-check.test.js`.

On a desk it installs nothing: both versions run as the portable
executable (or the unpacked build) with `--user-data-dir` in the temporary
folder, so the machine's own copy and profile are not touched.

```
node tools/upgrade-test.js
node tools/upgrade-test.js --from v2.0.0
node tools/upgrade-test.js --from <older portable .exe> --to "apps\desktop\release\win-unpacked\Rutba Office.exe"
```

With no `--to` it takes this checkout's build of package.json's version
(`npm run dist:win`, or `electron-builder --dir` for the unpacked copy);
with no `--from`, the newest GitHub release older than that. A tag's file
is fetched from its GitHub release into a cache in the temporary folder.
`--no-feed` leaves out the feed for a run with no network.

On GitHub, `.github/workflows/upgrade.yml` runs when a release is published
and from the Actions tab for any two tags: once installed (the
older installer, then the newer one over it, and the registry must show
one copy, of the newer version) and once portable, each opening a document
in the older version so it writes the recent list itself. A release made
before the upgrade report existed cannot be the newer side: the report is
the newer version's to write.

Two more jobs race the app's own closing, as an update does: the updater
starts the new installer the moment the app begins to quit, and 1.51 to
2.0 stopped there on "could not be removed ... (Uninstaller error 2)".
`--race close` starts the older copy again, runs the newer installer
silently while it is open (`/S --updated`, as the updater runs it) and
closes the copy 2 to 5 seconds later: the install must succeed and the
report follow. `--race never-close` leaves the copy open: the installer closes it
itself (its windows are asked after 5 seconds, what is left is ended after
10), must end with 0 and not hang, and must leave one copy, of the newer
version. Both need a newer
installer with the retry in `resources/installer.nsh`.
