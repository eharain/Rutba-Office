# Reproducing a build

How to build a release of Rutba Office again from its source, and how to
tell whether what came out is what was published.

## What a release is built from

- **The commit.** Each release is the tag `v<version>`, and its software
  bill of materials (`Rutba-Office-<version>-sbom.cdx.json`, one of the
  release's files) names the commit in `metadata.component.properties`
  (`rutba:commit`).
- **The packages.** `package-lock.json` fixes every npm package and its
  SHA-512; `npm ci` installs exactly those. Electron (43.4.1 at 2.0.0) and
  electron-builder come from the lock too.
- **The vendored files.** pdf.js, its fonts and decoders
  (`apps/desktop/vendor/pdfjs`) and the text recognition engine and its
  models (`packages/ocr/vendor`) are in the repository, not fetched. Their
  versions are in the `VERSION` file beside them, and
  `tools/vendor-pdfjs.mjs` and `tools/vendor-ocr.mjs` say where they came
  from.
- **The machine.** The release workflow (`.github/workflows/release.yml`)
  builds on GitHub's `windows-latest` with Node 22.

## Building it again

On Windows, with Node 22 and git:

```
git clone https://github.com/eharain/Rutba-Office.git
cd Rutba-Office
git checkout v<version>
npm ci
npm run dist:win
node tools/sbom.mjs --out rebuilt.cdx.json
```

`npm run dist:win` bundles the renderer (`apps/desktop/build/bundle.js`)
and packs the installer and the portable executable into
`apps/desktop/release`. electron-builder fetches the Electron and NSIS
binaries it packs into its own cache the first time; their versions are
fixed by the lock.

## Telling whether it is the same

The installers themselves are not identical byte for byte from one build
to the next: the NSIS installer and the portable executable carry the time
they were made, and the executable's resources are edited when it is
packed. What is compared instead is what they carry:

```
node tools/sbom.mjs --compare Rutba-Office-<version>-sbom.cdx.json rebuilt.cdx.json
```

The bill lists every file of the application's own (`main/`, the renderer
bundle in `build/out/`, the icons) with its SHA-256, every npm package
with its version and the lock's SHA-512, and every vendored file with its
SHA-256. "the same files, packages and versions" means the rebuild packs
the same program; anything else is listed by name. The bill itself is the
same for the same commit: its time is the commit's (or `SOURCE_DATE_EPOCH`
when that is set), its serial number is made from its content, and its
lists are sorted.

To look inside an installer's archive directly, unpack the build
(`apps/desktop/release/win-unpacked/resources/app.asar`) with
`npx @electron/asar extract app.asar <folder>` and compare folders.

## What is not covered

- Electron's own binaries are taken as electron-builder fetches them for
  the version in the lock; they are named in the bill by version and the
  lock's hash of the npm package, not rebuilt.
- The installers are unsigned, so there is no signature to compare.
- A build on another operating system packs the same files for a
  different platform; compare bills made on the same one.
