# Locker file formats

Every file the Locker writes says which version of the format it is. This page
is the record of what each version means, so that a file locked years ago can
still be opened.

## The rules

1. A released format never changes. Any change to the layout, an algorithm or
   a context string is a new version number.
2. The code that reads an old version stays in Rutba Office for good.
3. A version this copy does not know is reported as "needs a newer Rutba
   Office" (error code `newer`). It is never reported as damage, never written
   over and never converted.
4. `tests/fixtures/locker-format-1` holds real files of version 1. They are
   never regenerated. `tests/locker-format.test.js` must open them in every
   later version.

To add version 2: raise `FORMAT_VERSION` in `format.js`, keep the version 1
path as a branch on the number that was read, add a section below, and add
`tests/fixtures/locker-format-2`.

## Where the version is written

| What | Where | Version 1 |
| :-- | :-- | :-- |
| A locked file `*.rlock`, a shared file, a list `*.rlist`, an earlier version in `.rlocker-old/` | first 8 bytes | `RLOCKV1\n` |
| The same file's sealed details (frame 0) | field `v` | `1` |
| A piece of a large file `*.rpiece` | first 8 bytes | `RPIECEV1` |
| A large file's `item.json` | field `v` | `1` |
| A large file's sealed list body | field `v` | `1` |
| A folder's `.rutbalocker-keys` | field `v` | `1` |
| The profile's `keyring.json` | field `v` | `1` |
| A safety copy `safety/<id>.rsafe`, in the profile | first 8 bytes, and its details' `v` | `RLOCKV1\n`, `1` |
| Its archive's first entry, the list of files | field `v` | `1` |
| The profile's `safety/key.json` and `safety/copies.json` | field `v` | `1` |
| Every key derivation and every associated-data string | prefix | `rutba-locker:v1:` |

The profile's key store is the `@tech-style/secure-keystore` library's own
format, and the library records its own version.

## Version 1

Algorithms, all from `node:crypto`: AES-256-GCM, HKDF-SHA256, scrypt, SHA-256,
HMAC-SHA256.

### Keys

- The locker's root key is 32 random bytes, stored only wrapped.
- A folder's key is 32 random bytes.
- A password becomes a key with scrypt. Its salt, N, r and p are stored beside
  what it wraps. New passwords use N = 2^17, r = 8, p = 1. The text is
  normalised to NFKC and encoded as UTF-8 first.
- A recovery key (32 random bytes, written in a 32-letter alphabet) or a
  Windows Hello secret becomes a key with HKDF-SHA256, salt as stored, info
  `rutba-locker:v1:recovery` or `rutba-locker:v1:hello`.
- The Windows Hello secret is the SHA-256 of Windows' RSA-2048 PKCS#1 SHA-256
  signature of `rutba-locker:v1:hello-challenge`, by a key named
  `RutbaOfficeLocker-<id>`. It exists on that computer only.
- A wrapped key is `{ iv, tag, data }` in hex: AES-256-GCM, associated data
  `rutba-locker:v1:wrap:<context>`. The context is `password`, `recovery` or
  `hello` in the keyring and `folder:<id>` in a folder's key file.

### A locked file

```
offset  size
0       8    "RLOCKV1\n"
8       4    chunk size, u32 BE, always 65536
12      1    mode: 0 the folder's key, 1 a password of its own (a shared file)
13      3    zero
16      16   the file's id
32      32   the file's salt
64      16   the folder's id (zero for a shared file)
80      32   scrypt salt (a shared file only)
112     12   scrypt N, r, p, u32 BE each (a shared file only)
124     4    zero
128          frame 0: details, 4096 bytes of JSON padded with zero bytes
             frames 1..n: the file, 64 KB each, the last one shorter
```

A frame on disk is `epoch(7) || ciphertext || tag(16)`. Its nonce is
`epoch(7) || index(u32 BE) || last(1)`. Its associated data is
`rutba-locker:v1:chunk || SHA-256(header) || index(u64 BE) || last(1)`.

The file's key is HKDF-SHA256 of the folder's key (or, for a shared file, the
scrypt key of its password), the file's salt, and info
`rutba-locker:v1:file:<file id in hex>`.

Frame 0 holds `v`, `name`, `size`, `sha256` of the whole plain file, `frames`
and `modified`.

### A large file (`<name>.rlock.d`)

- `item.json`: `{ v, fileId, salt }`, no secrets.
- `p/<name>.rpiece`: `"RPIECEV1" || nonce(12) || ciphertext || tag(16)`. The
  name is the first 40 hex digits of an HMAC-SHA256 of the plain piece under
  an HKDF key with info `rutba-locker:v1:piece:name:<file id>`. The piece's key
  is HKDF with info `rutba-locker:v1:piece:<file id>:<name>`, which is also its
  associated data.
- `list-<seq>-<rand>.rlist`: a locked file as above, under the folder's key,
  holding JSON `{ v, name, size, sha256, pieces: [{ name, size }], modified }`.

### A folder's key file (`.rutbalocker-keys`)

`{ v, id, unlockers }`. Each unlocker wraps the folder's key: `kind: "password"`
with its `kdf`, or `kind: "recovery"` with its `salt`.

### The keyring (`keyring.json`, in the profile)

`{ v, made, unlockers }`. Each unlocker wraps the root key: `kind` is
`password` with `kdf`, `recovery` with `salt`, or `hello` with `id` and `salt`.

### A safety copy (`safety/<id>.rsafe`, in the profile)

Added after 2.25.0. It is a locked file exactly as above, so nothing in this
version's layout changed. Only what it holds and the key it is under are new.

- The key is in `safety/key.json`: `{ v, id, salt, by, sealed, made }`, made
  once per Office profile. The key itself is 32 random bytes, never written
  plain: `sealed` is it sealed by the system, in base64, and `by` says how.
  - `by: "system"`, on every system: Electron's `safeStorage.encryptString`
    of the key in base64, opened with `safeStorage.decryptString`. Electron
    uses DPAPI for this account on Windows, the Keychain on macOS and the
    secret service (libsecret or KWallet) on Linux. On Windows, Electron
    keeps its own key for this, sealed by DPAPI, in `Local State` in Office's
    user data folder: if that file is lost, the copies cannot be opened any
    more.
  - Where `safeStorage.isEncryptionAvailable()` is false, or Linux has no
    keyring and Electron would fall back to a fixed key (`basic_text`), no key
    is made and no copy is made. The key is never written plain.
  - No other `by` is read: a key sealed some other way is reported as sealed
    on another kind of system.
- Where the Office profile has a password, the whole of `key.json` is also
  sealed by the profile's vault (`packages/office-shell/src/profiles.js`,
  `seal`), and sealed again whenever that password changes. Another Office
  profile on the same user account then cannot open the copies. Without a
  profile password it can.
- The file's root key is HKDF-SHA256 of the key, salt as stored, info
  `rutba-locker:v1:safety`. The header's folder id is the key's `id`. The key
  is not wrapped by the Locker password: anyone who opens it as above can open
  a safety copy.
- What is sealed is a ZIP archive: every entry stored (method 0), its name in
  UTF-8 (flag bit 11), and its sizes in a ZIP64 field (header 0x0001) in both
  the local header and the central directory, which ends with a ZIP64 end
  record, its locator and an ordinary end record.
- The first entry is `.rutba-safety-copy.json`: `{ v, folder, made, files:
  [{ rel, size, sha256, modified }] }`. The files follow in that order, each
  under its path in the folder with `/` between the parts.
- Frame 0 holds `v`, `name`, `kind: "safety"`, `made`, `size`, `sha256`,
  `frames` and `modified`, as any locked file does.
- `safety/copies.json` is `{ v, copies: [{ id, folderId, root, name, file,
  made, files, bytes, size }] }`: no file names, only the folder's.
- A `.rl-<24 hex>.part` in `safety/` is a copy being written. Left by a stop
  midway, it is wiped before the next copy is made, as is a `.rsafe` that
  `copies.json` does not list, once either is 15 minutes old (a newer one may
  be another program's copy in progress). A listed copy whose folder still
  asks for one is used only if it opens; otherwise a new copy replaces it.

## An opened file's plain copy (not a format, and never in a locked folder)

While a program has a locked file open, its plain copy is at
`<profile>/locker/open/<owner>-<random>/<Name.ext>`, in Office's own profile
folder. The profile is the person's own on every system, so the Locker sets no
permissions on it: the folders are made 0700 where the system keeps such modes,
and on Windows they keep what the profile hands down. `<owner>` names the
locker that made the folder. `open.json`, beside `state.json`, lists the
copies before any is written, so a crash can be cleaned up after. A copy is
written back through the locked file's own save, and removed when the program
lets go, when the Locker locks, and at the next unlock after a crash.

2.24.0 and the first 2.25.0 builds kept the copy in a hidden folder beside the
file, `.rutba-open/<owner>-<random>/`. Since 2.25.0 such a folder is only
cleared: at the next unlock, once nothing open is in it (one with no owner in
its name waits a day, as it may be another computer's). A folder record may
still hold `openIn`; it is not read.
