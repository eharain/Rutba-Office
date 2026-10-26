# The Data Files format (.rbase)

**Status.** Specification version 7. The number is the design version a
file states in `rutba_meta` (key `version`), and this text describes design
version 7 exactly as Rutba Office 2.19.0 writes it. It was first published
on 2026-10-26, as version 6; section 7 says what version 7 changed.

**Licence.** This text is published under the Creative Commons Attribution
4.0 International licence (CC BY 4.0,
<https://creativecommons.org/licenses/by/4.0/>), and so are the sample
files in `docs/data-files-format/samples/`. Anyone may implement the format,
in any program and under any licence, free of charge and without asking.
Tech Style Ltd (tech-style.co), which publishes Rutba Office, makes no
patent claim against any implementation of the format. The checker,
`tools/check-rbase.mjs`, is part of Rutba Office and is under the
repository's own licence; the format does not depend on it.

**Changes.** The format changes only with a new design version, and each
new design version comes with a new version of this text, made in the open
in this repository. The history of this file is the history of the format.

The words *must*, *must not*, *should* and *may* are used in their ordinary
sense: *must* is a rule a file or a program has to follow to follow this
specification, *should* is advice, *may* is a permission.

## Contents

1. [The file](#1-the-file)
2. [Where the data lives](#2-where-the-data-lives)
3. [The design tables](#3-the-design-tables)
4. [The layouts kept as JSON](#4-the-layouts-kept-as-json)
5. [Relationships](#5-relationships)
6. [The history of changes](#6-the-history-of-changes)
7. [Design versions and upgrading](#7-design-versions-and-upgrading)
8. [Protected databases](#8-protected-databases)
9. [Linked tables, synced copies and other sources](#9-linked-tables-synced-copies-and-other-sources)
10. [Rules for readers and writers](#10-rules-for-readers-and-writers)
11. [Where Rutba Office departs from this text](#11-where-rutba-office-departs-from-this-text)
12. [Sample files and the checker](#12-sample-files-and-the-checker)

## 1. The file

A Data Files database is one SQLite 3 database file
(<https://www.sqlite.org/fileformat.html>). Nothing else is part of it:
there is no wrapper, no second file and no compression.

- **Extension.** `.rbase`. No media type is registered; a program that
  needs one should use `application/vnd.sqlite3`.
- **Recognising one.** A file is a Data Files database when it is an SQLite
  database (its first 16 bytes are `SQLite format 3` followed by a zero
  byte) that has the tables `rutba_meta`, `rutba_tables` and `rutba_fields`,
  and `rutba_meta` has a row whose `key` is `version`. The extension is a
  hint, not proof. An SQLite database with tables of its own and no
  `rutba_meta` is another program's (section 10, rule 1).
- **application_id and user_version.** Rutba Office sets neither: both are
  0 in every file it writes. A reader must not rely on them, and a writer
  should leave them at 0, so that the design version is stated in one place
  only (`rutba_meta`).
- **Page size.** Not part of the format. Rutba Office uses SQLite's default
  (4096 bytes). Any page size SQLite allows is allowed.
- **Text encoding.** Rutba Office writes UTF-8. Any encoding SQLite supports
  is allowed; values are read through SQLite, which converts them.
- **Journal mode.** Not part of the format. Rutba Office switches a file to
  WAL mode each time it opens it (`PRAGMA journal_mode = WAL`), so a file it
  wrote usually says so in its header (bytes 18 and 19 are 2), and while the
  file is open there are `-wal` and `-shm` files beside it. Those belong to
  the open file: a program that copies a database should copy it with
  `VACUUM INTO` or after a checkpoint, never by copying the main file alone
  while it is open. The sample files are in rollback mode (bytes 18 and 19
  are 1).
- **SQLite version.** The schema uses only long-standing SQLite syntax:
  `INTEGER PRIMARY KEY AUTOINCREMENT`, `CHECK`, `FOREIGN KEY ... ON UPDATE
  ... ON DELETE`, `GLOB` and the built-in functions `typeof`, `length` and
  `round`. Any SQLite 3.7.0 or later reads it (3.7.0 is the first with WAL).
  A writer that changes the design the way Rutba Office does needs SQLite
  3.25.0 or later (`ALTER TABLE ... RENAME COLUMN`). Rutba Office 2.19.0
  writes with the SQLite built into its runtime (3.50 at the time of
  writing).
- **Foreign keys.** Relationships are real SQLite foreign keys (section 5).
  SQLite enforces them only on a connection that has run `PRAGMA
  foreign_keys = ON`, and a writer must run it.
- **Nothing of Rutba Office's own is needed to read the file.** The file
  holds no views and no triggers, and no `CHECK` calls a function that
  SQLite does not have. Rutba Office makes TEMP tables and TEMP views on its
  own connection (for linked tables and for roles, sections 8 and 9); they
  are never written to the file.
- **Names the file keeps for itself.** Every table whose name starts with
  `rutba_` (in any case) belongs to the design, and so does every name
  starting with `sqlite_` (SQLite's own). A person's table or field may not
  have such a name. `sqlite_sequence` holds the AutoNumber counters
  (section 2.2).

## 2. Where the data lives

Each table of the database is an ordinary SQLite table of the same name,
with one column for each field, so any SQLite tool reads and writes the
records. The design tables (`rutba_*`, section 3) say what the tables and
fields are: their order, types, captions and relationships. A table is part
of the database when it has a row in `rutba_tables`. A table of the file
that has no row there is not part of the design (section 11).

### 2.1 Names

The names of tables, fields, queries, forms, reports, print templates,
roles, people and links follow one rule: trimmed of spaces at each end,
1 to 64 characters, no double quotation mark (`"`), no control character
(U+0000 to U+001F), and not starting with `rutba_` or `sqlite_` in any
case. Two tables may not have names that differ only in case, nor may two
fields of one table. A saved query may not have the name of a table, nor a
linked table the name of a query (compared without case). Names are kept
as typed, and every identifier in SQL is written in double quotes.

### 2.2 Field types

`rutba_fields.type` holds one of nine words. Each is a column of the
declared type shown, with a `CHECK` that keeps what SQLite would not keep
by itself. The table shows the column exactly as Rutba Office writes it,
for a field named `F` (written `"F"`):

| Type | Column | Values |
|---|---|---|
| `autonumber` | `"F" INTEGER PRIMARY KEY AUTOINCREMENT` | A whole number given by SQLite when a record is added. The column is the table's rowid. |
| `text` | `"F" TEXT CHECK ("F" IS NULL OR (typeof("F") = 'text' AND length("F") <= N))` | Text of at most N characters. N is the field's `size`, 255 when it has none, and is kept between 1 and 4000. |
| `longtext` | `"F" TEXT CHECK ("F" IS NULL OR (typeof("F") = 'text'))` | Text of any length. Plain text: no formatting is kept. |
| `number` | `"F" REAL CHECK ("F" IS NULL OR (typeof("F") IN ('real', 'integer')))` | A number, kept as an IEEE 754 double. |
| `integer` | `"F" INTEGER CHECK ("F" IS NULL OR (typeof("F") = 'integer'))` | A whole number (64-bit). |
| `currency` | `"F" REAL CHECK ("F" IS NULL OR (typeof("F") IN ('real', 'integer') AND round("F", 4) = "F"))` | An amount of money with at most 4 decimal places, kept as a double. No currency is stored. |
| `date` | `"F" TEXT CHECK ("F" IS NULL OR ("F" GLOB '[0-9][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9]'))` | A day as `YYYY-MM-DD`. |
| `datetime` | `"F" TEXT CHECK ("F" IS NULL OR ("F" GLOB '[0-9][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9]*'))` | A day and a time as `YYYY-MM-DDTHH:MM:SS`. |
| `yesno` | `"F" INTEGER NOT NULL DEFAULT 0 CHECK ("F" IN (0, 1))` | 1 for Yes, 0 for No. `DEFAULT 1` when the field's default is Yes. |

After the `CHECK`, every type but `autonumber` and `yesno` gains
` NOT NULL` when the field is required, ` UNIQUE` when it is unique, and
` DEFAULT <literal>` when it has a default (text in single quotes with its
quotes doubled, a number as written, true and false as 1 and 0). A
`yesno` is never empty, so required, unique and a default of No change
nothing in its column; its default of Yes is its `DEFAULT 1`.

What the types mean for a reader and a writer:

- **AutoNumber.** A table made by Rutba Office has exactly one AutoNumber,
  made with the table and kept for its life: called `ID`, or `<Table> ID`
  when the table has another field called ID. Its counter is the row of
  `sqlite_sequence` for the table. A number given once is never given again,
  even after its record is deleted or the table is made again (Rutba Office
  copies the counter across when it rebuilds a table). A writer must not
  change an AutoNumber's value and must not reuse a number.
- **Money.** A currency field is a binary double rounded to 4 places when it
  is stored, not a decimal type. A reader should show a value to 15
  significant digits (Rutba Office shows 31.240000000000002 as 31.24) and
  should add amounts the same way. The currency symbol is the reader's
  choice; the file does not say which currency an amount is in.
- **Dates and times.** A date is the text `YYYY-MM-DD`. A date and time is
  the text `YYYY-MM-DDTHH:MM:SS`, always with seconds, with no fraction and
  no time zone: it is the time as a person entered it, a wall-clock time.
  The `CHECK` patterns are looser than this (they accept a month of 19, and
  anything after the seconds of a datetime); a writer must write real
  calendar dates and exactly these forms.
- **Yes/No.** Stored as the integer 1 or 0, never empty.
- **Empty values.** An empty value is SQL `NULL`. Rutba Office stores an
  empty text as `NULL`, not as `''`; a reader should treat both as empty.
- **Pictures and attachments.** Design version 7 has no field type for a
  picture, a file or an attachment, and no table holds files. A reader must
  not expect one.
- **Lookups.** There is no lookup type. A field that points at another
  table's record is an ordinary field (an `integer`, as Rutba Office makes
  it) that is the "from" side of a relationship (section 5). Forms show it
  as a list of the other table's records, and reports and print templates
  show the words that name the record it points at (section 3.1,
  `shown`).

## 3. The design tables

The design is kept in tables whose names start with `rutba_`. Each is made
with `CREATE TABLE IF NOT EXISTS` exactly as shown under its heading below.
The "since" column of this list is the design version that added the table
(section 7). A file of design version V must have every table of its list
whose "since" is V or less.

| Table | Since | Holds |
|---|---|---|
| `rutba_meta` | 1 | The design version and a few settings, as keys and values. |
| `rutba_tables` | 1 | The tables, their captions and order. |
| `rutba_fields` | 1 | Every field of every table: type, size, required, unique, default, caption, order. |
| `rutba_relations` | 1 | The relationships between tables. |
| `rutba_steps` | 1 | The numbered history of changes to the design. |
| `rutba_queries` | 2 | Saved queries, each a JSON design. |
| `rutba_forms` | 2 | Forms, each a JSON layout. |
| `rutba_reports` | 3 | Reports, each a JSON layout. |
| `rutba_documents` | 4 | Print templates (invoices, receipts, payslips), each a JSON layout. |
| `rutba_roles` | 5 | Roles and what each may see and do. |
| `rutba_people` | 5 | The people who may sign in to a protected database. |
| `rutba_trail` | 5 | Who added, changed and deleted which records. |
| `rutba_links` | 6 | Database engines and other sources the database is linked to. |
| `rutba_linked` | 6 | Tables linked to an engine or read from a source. |
| `rutba_sync` | 7 | For each record of a synced copy, its state against the engine. |
| `rutba_clashes` | 7 | Changes made here that the engine's version replaced. |

Every time below is an ISO 8601 time in UTC with milliseconds, as
JavaScript's `toISOString()` writes it: `2026-10-26T09:39:51.157Z`.
Booleans are the integers 1 and 0.

### rutba_meta

```sql
CREATE TABLE IF NOT EXISTS rutba_meta (key TEXT PRIMARY KEY, value TEXT);
```

| Column | Meaning |
|---|---|
| `key` | The name of a setting. |
| `value` | Its value, as text. Several values are JSON. |

#### Keys of rutba_meta

| Key | Value |
|---|---|
| `version` | The design version, a whole number written in decimal: `7`. Required. |
| `protected` | `1` while the database is protected (section 8), `0` after the protection was taken away. Absent means not protected. |
| `shown` | JSON object: for a table, the field, or list of fields, that names its records in lists, reports and print templates. `{"Customers":["Name"],"People":["First Name","Last Name"]}`. A table not in it is named by its data (below). |
| `pinned` | JSON array of the items pinned to the top of the side panel: `[{"kind":"table","name":"Orders"}]`, `kind` one of `table`, `query`, `form`, `report`, `document`. |
| `folded` | JSON array of the side panel's sections folded away: any of `tables`, `queries`, `forms`, `reports`, `documents`. |
| `origin` | JSON object, present when the database was made from one of Rutba Office's ready-made databases: `{"kind":"template"\|"sample","id":"<id>","set":"<template set>"\|null,"made":"<time>"}`, with `updated` in place of or beside `made` after a newer design was applied. |
| `origin.design` | JSON object of fingerprints of the ready-made design as made: `{"tables":{"<table>":["<field>", ...]},"queries":{"<name>":"<print>"},"forms":{...},"reports":{...},"documents":{...}}`. A print is the first 20 hexadecimal digits of the SHA-256 of the item's stored JSON text. Rutba Office compares them to offer a newer ready-made design without replacing what the person changed. Other programs may ignore both keys, and must keep them. |

When `shown` does not name a table's fields, Rutba Office works out what
names a record from the data: a field called Name or Title, a text field
whose name ends in Name, a first and last name together, then the first
text field whose values are different in at least 9 records of 10; and when
nothing names one record each, the table's singular noun with the
AutoNumber and the first date ("Sale 17, 21/10/2026"). This is a choice of
display, not part of the format: another program may name records its own
way.

### rutba_tables

```sql
CREATE TABLE IF NOT EXISTS rutba_tables (name TEXT PRIMARY KEY, caption TEXT, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The table's name, which is also the name of its SQLite table. |
| `caption` | Words to show for the table instead of its name, or `NULL`. |
| `position` | The table's place in the list, from 0. Tables are listed by `position`, then `name`. |

A table linked to an engine or read from a source is not in
`rutba_tables`: it is in `rutba_linked` (section 9).

### rutba_fields

```sql
CREATE TABLE IF NOT EXISTS rutba_fields (tbl TEXT NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL, size INTEGER, required INTEGER NOT NULL DEFAULT 0,
  uniq INTEGER NOT NULL DEFAULT 0, dflt TEXT, caption TEXT, position INTEGER NOT NULL, PRIMARY KEY (tbl, name));
```

| Column | Meaning |
|---|---|
| `tbl` | The table: a `name` in `rutba_tables`, or a `tbl` in `rutba_linked`. |
| `name` | The field's name, which is also its column's name. |
| `type` | One of the nine types of section 2.2. |
| `size` | For a `text` field, its longest length in characters; `NULL` means 255. `NULL` for the other types. |
| `required` | 1 when the field must have a value (`NOT NULL`), else 0. |
| `uniq` | 1 when no two records may have the same value (`UNIQUE`), else 0. |
| `dflt` | The default as JSON (`"UK"`, `0`, `true`), or `NULL` for none. The column's `DEFAULT` clause holds the same value. |
| `caption` | A label to show for the field on forms and reports, or `NULL` for one made from its name. |
| `position` | The field's place in its table, from 0, in the order of the table's columns. |

The rows of a table's own fields describe its SQLite table: every column of
the SQLite table has a row here, and every row here a column there, with the
column made as section 2.2 says.

### rutba_relations

```sql
CREATE TABLE IF NOT EXISTS rutba_relations (id INTEGER PRIMARY KEY, from_tbl TEXT NOT NULL, from_field TEXT NOT NULL, to_tbl TEXT NOT NULL,
  to_field TEXT NOT NULL, cascade_update INTEGER NOT NULL DEFAULT 1, cascade_delete INTEGER NOT NULL DEFAULT 0);
```

| Column | Meaning |
|---|---|
| `id` | The relationship's number. |
| `from_tbl` | The table on the "many" side: the one whose field points (Orders). |
| `from_field` | Its field that points (Orders.Customer). |
| `to_tbl` | The table pointed at (Customers). |
| `to_field` | Its field pointed at, which must be its AutoNumber or a unique field (Customers.ID). |
| `cascade_update` | 1: when the value pointed at changes, the records pointing at it change with it (`ON UPDATE CASCADE`). 0: such a change is refused (`ON UPDATE RESTRICT`). |
| `cascade_delete` | 1: deleting a record deletes the records pointing at it (`ON DELETE CASCADE`). 0: a record others point at cannot be deleted (`ON DELETE RESTRICT`). |

Section 5 says how a relationship is kept in the tables themselves.

### rutba_steps

```sql
CREATE TABLE IF NOT EXISTS rutba_steps (n INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, label TEXT NOT NULL, detail TEXT);
```

| Column | Meaning |
|---|---|
| `n` | The step's number, from 1, never reused. |
| `at` | When the step was made (UTC). |
| `label` | The step in English words: `Create table Orders`, `Query Overdue kept`. |
| `detail` | JSON object saying what changed, or `NULL`. |

Section 6 lists the steps and says how a reader treats them.

### rutba_queries

```sql
CREATE TABLE IF NOT EXISTS rutba_queries (name TEXT PRIMARY KEY, design TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The query's name. |
| `design` | The query's design as JSON (section 4.1). |
| `position` | Its place in the list, from 0. Listed by `position`, then `name`. |

### rutba_forms

```sql
CREATE TABLE IF NOT EXISTS rutba_forms (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The form's name. |
| `layout` | The form's layout as JSON (section 4.2). Its `name` is the row's `name`. |
| `position` | Its place in the list, from 0. |

### rutba_reports

```sql
CREATE TABLE IF NOT EXISTS rutba_reports (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The report's name. |
| `layout` | The report's layout as JSON (section 4.3). Its `name` is the row's `name`. |
| `position` | Its place in the list, from 0. |

### rutba_documents

```sql
CREATE TABLE IF NOT EXISTS rutba_documents (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The print template's name. |
| `layout` | The print template's layout as JSON (section 4.4). Its `name` is the row's `name` and its `kind` is `document`. |
| `position` | Its place in the list, from 0. |

### rutba_roles

```sql
CREATE TABLE IF NOT EXISTS rutba_roles (name TEXT PRIMARY KEY COLLATE NOCASE, rights TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `name` | The role's name, unique without regard to case. `Admin` is built in. |
| `rights` | What the role may see and do, as JSON (section 4.5). |
| `position` | Its place in the list; Admin is 0. |

### rutba_people

```sql
CREATE TABLE IF NOT EXISTS rutba_people (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE COLLATE NOCASE, role TEXT NOT NULL,
  hash TEXT, account TEXT, created TEXT NOT NULL);
```

| Column | Meaning |
|---|---|
| `id` | The person's number, never reused. `rutba_trail.person` refers to it. |
| `name` | The name they sign in with, unique without regard to case. |
| `role` | The `name` of their role in `rutba_roles`. |
| `hash` | Their password, kept as section 8.2 says. Never the password itself. |
| `account` | Kept for a later version. Rutba Office 2.19.0 never writes it, so it is `NULL`. |
| `created` | When they were added (UTC). |

### rutba_trail

```sql
CREATE TABLE IF NOT EXISTS rutba_trail (n INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, person INTEGER, tbl TEXT NOT NULL, row INTEGER NOT NULL,
  action TEXT NOT NULL, detail TEXT);
CREATE INDEX IF NOT EXISTS rutba_trail_row ON rutba_trail (tbl, row);
```

| Column | Meaning |
|---|---|
| `n` | The entry's number. |
| `at` | When (UTC). |
| `person` | The `id` in `rutba_people` of who did it. It may name a person since removed. |
| `tbl` | The table. |
| `row` | The record's rowid in that table. |
| `action` | `add`, `change` or `delete`. |
| `detail` | For `change`, JSON `{"fields":["<field>", ...]}` naming the fields written; otherwise `NULL`. |

An entry is written for each record added, changed or deleted while the
database is protected and someone is signed in. A role that sees only its
own records (section 4.5) sees those whose `add` entry names its person, so
a writer of a protected database must keep the trail.

### rutba_links

```sql
CREATE TABLE IF NOT EXISTS rutba_links (id TEXT PRIMARY KEY, name TEXT NOT NULL UNIQUE COLLATE NOCASE, engine TEXT NOT NULL, settings TEXT NOT NULL, position INTEGER);
```

| Column | Meaning |
|---|---|
| `id` | The link's identifier. Rutba Office makes `link-` followed by the time in base 36, a hyphen and six random base-36 characters (`link-mgx2k1qz-4f8a0b`); any unique text is allowed. |
| `name` | A name for the link, unique without regard to case. |
| `engine` | What it reaches: `postgres` (PostgreSQL), `mysql` (MySQL or MariaDB), or a source: `file` (a CSV file or a spreadsheet), `rbase` (another Data Files database), `json` (a web address that answers JSON). |
| `settings` | How to reach it, as JSON (section 9.1). Never a password or other secret. |
| `position` | Its place in the list. |

### rutba_linked

```sql
CREATE TABLE IF NOT EXISTS rutba_linked (tbl TEXT PRIMARY KEY COLLATE NOCASE, link TEXT NOT NULL, remote TEXT NOT NULL, key TEXT, caption TEXT, position INTEGER);
-- part of design version 7 (Rutba Office 2.18.0 added them, still as version 6, the first time a table was kept as a copy):
ALTER TABLE rutba_linked ADD COLUMN mode TEXT NOT NULL DEFAULT 'live';
ALTER TABLE rutba_linked ADD COLUMN stamp TEXT;
ALTER TABLE rutba_linked ADD COLUMN synced TEXT;
```

| Column | Meaning |
|---|---|
| `tbl` | The table's name in this database. Not also in `rutba_tables`. |
| `link` | The `id` of its link in `rutba_links`. |
| `remote` | Which table of the engine or which part of the source, as JSON (section 9.1). |
| `key` | The engine's primary key field, or `NULL` when it has none: then the table can be read, not changed. |
| `caption` | Words to show instead of the name, or `NULL`. |
| `position` | Its place in the list. |
| `mode` | `live`: the records stay in the engine and are read each time. `copy`: the records are kept in this file and synced. When the column is absent (a version 6 file), every linked table is `live`. A source's table is always `copy`. |
| `stamp` | For a copy, a field the engine changes with every change to a record (an updated-at time), so only changed records are read; `NULL` to read the whole table. |
| `synced` | For a copy, when it last caught up with its engine or source (UTC), or `NULL`. |

The fields of a linked table are rows of `rutba_fields`, as the engine or
source last described them.

### rutba_sync

```sql
CREATE TABLE IF NOT EXISTS rutba_sync (tbl TEXT NOT NULL COLLATE NOCASE, rid INTEGER NOT NULL, kj TEXT, hash TEXT, state TEXT, n INTEGER NOT NULL DEFAULT 0, since TEXT, PRIMARY KEY (tbl, rid));
CREATE INDEX IF NOT EXISTS rutba_sync_key ON rutba_sync (tbl, kj);
```

| Column | Meaning |
|---|---|
| `tbl` | A table of `rutba_linked` whose `mode` is `copy`. |
| `rid` | The record's rowid in the copy. A record deleted here and not yet deleted in the engine keeps its entry under a negative number, one below the lowest `rid` of the table and below 0, so a record added later at its rowid has an entry of its own. |
| `kj` | The record's key in the engine, as JSON text (`42`, `"A-17"`); `null` for a record added here that the engine has not numbered yet. |
| `hash` | The hash of the engine's version of the record last seen (section 9.3), or `NULL` when it is to be read again. |
| `state` | `NULL`: the record is as the engine had it. `added`, `changed` or `deleted`: changed here and waiting to be written to the engine. |
| `n` | How many times the record was changed here; a change made while an earlier one is on its way is seen by a different `n`. |
| `since` | When the first change still waiting was made (UTC), or `NULL`. |

### rutba_clashes

```sql
CREATE TABLE IF NOT EXISTS rutba_clashes (id INTEGER PRIMARY KEY, tbl TEXT NOT NULL COLLATE NOCASE, kj TEXT, at TEXT NOT NULL, action TEXT NOT NULL, mine TEXT, theirs TEXT, why TEXT);
```

| Column | Meaning |
|---|---|
| `id` | The clash's number. |
| `tbl` | The copy's table. |
| `kj` | The record's key in the engine, as JSON text, or `NULL`. |
| `at` | When the engine's version replaced the one from here (UTC). |
| `action` | `changed`: changed here and in the engine; the engine's version was kept. `gone`: changed here, deleted in the engine; it was deleted here too. `deleted`: deleted here, changed in the engine; the engine's version came back. `refused`: the engine would not take the change. |
| `mine` | The version from here, a JSON object of field to value, or `NULL`. |
| `theirs` | The engine's version, a JSON object, or `NULL`. |
| `why` | For `refused`, the engine's words; otherwise `NULL`. |

A clash stays until a person dismisses it or puts their version back.

## 4. The layouts kept as JSON

Queries, forms, reports, print templates and roles are kept as JSON text:
UTF-8, an object at the top. The keys below are those Rutba Office 2.19.0
writes and reads. A key marked "optional" may be absent. Any other key is
allowed and must be kept (section 10): a program that changes a layout
changes the keys it means to change and leaves the others as they are.

### 4.1 A query's design (rutba_queries.design)

A query is one of two things.

**Written in SQL.** `{"sql": "SELECT ..."}`: one `SELECT` statement (or
`WITH ... SELECT`), kept as written, trimmed and without a final semicolon.
It must be run with the database read-only (`PRAGMA query_only = ON`). It
may name any table. In a protected database only a role that may change the
design may run it, because it could read past a role's limits.

**Designed.** An object:

| Key | Value |
|---|---|
| `tables` | Array of table names (own or linked). The first is read; each other is joined to one before it by a relationship between them (an inner join on `from_field = to_field`, the first such relationship in `id` order). A table no relationship reaches makes the design invalid: tables are never crossed. |
| `columns` | Array of columns, in the order they are shown. |

A column is either a field, `{"table": "<table>", "field": "<field>"}`, the
table one of `tables`; or a calculation, `{"expr": "<expression>", "as":
"<heading>"}` (expressions below). Either has these keys besides:

| Key | Value |
|---|---|
| `as` | Optional. The column's heading; for a calculation with none, `Expr1`, `Expr2` and so on in order. |
| `show` | Optional, default `true`. `false`: used for criteria or sorting only. |
| `sort` | Optional: `null`, `"asc"` or `"desc"`. Sorting follows the columns' order. |
| `total` | Optional: `null`, or one of `group`, `sum`, `avg`, `count`, `min`, `max`, `where`. When any column has a total the query groups: `group` columns are grouped by, the aggregates are applied, and a `where` column is used for its criteria only and not shown. |
| `criteria` | Optional array of text, one cell for each Or row. The cells of one row are joined with AND across the columns, and the rows with OR. A criterion on an aggregated column applies to the groups (`HAVING`). |
| `decimals` | Optional whole number 0 to 6: the answer rounded to that many places. |

A heading is the field's name, the calculation's name, or for an aggregate
without `as`, `Sum of X`, `Average of X`, `Count of X`, `Min of X` or `Max
of X`. (Rutba Office 2.17 headed an average `Avg of X`; a report naming such
a heading is read as `Average of X`.)

**Criteria.** A cell is one of: `Is Null`; `Is Not Null`; `Like <pattern>`
and `Not Like <pattern>`, with `*` for any characters and `?` for one;
`In (<value>, ...)` and `Not In (...)`; `Between <a> And <b>` and `Not
Between ...`; an operator `=`, `<>`, `<`, `<=`, `>`, `>=` before a value;
`Not <criterion>`; or a value alone, meaning equal to it (like it when it
holds `*` or `?`). A value is `"text"` or `'text'` (the quote doubled
inside), a date `#2026-10-21#` or `#2026-10-21 14:30#`, a number, `True`,
`Yes`, `False` or `No` (1 and 0); a value in brackets names a field
(`<= [Reorder Level]`) and is worked out as an expression. Values must be
passed to SQL as parameters, never as SQL text.

**Expressions.** A calculation, as Access writes one: field names in
brackets, `[Quantity]` or `[Products].[Price]` (a bare name must belong to
one table of the query only); numbers; text in double or single quotes;
dates in `#...#`; `True`, `Yes`, `False`, `No`, `Null`; the operators
`+ - * /` (`/` always divides as decimals), `Mod`, `&` (joins text, an empty
side as nothing), `= <> < <= > >=`, `And`, `Or`, `Not` and parentheses; and
these functions, without regard to case: `Round(x[, places])`, `Nz(x[,
instead])`, `IIf(test, then, else)`, `Abs`, `Year`, `Month`, `Day`,
`Weekday` (1 is Sunday), `Date()`, `Now()`, `Format(x, pattern)`,
`DateDiff("d"|"m"|"yyyy", from, to)`, `Len`, `Left(x, n)`, `Right(x, n)`,
`Mid(x, start[, n])`, `UCase`, `LCase`, `Trim`. A `Format` pattern is
`000` (padded with noughts), `0.00`, `#,##0` or `#,##0.00`, or a date
pattern of `yyyy yy mmmm mmm mm m dddd ddd dd d hh h nn ss` with any other
characters kept as written (English month and day names). Nothing else is
an expression, so an expression can compute a value and do nothing else.

### 4.2 A form's layout (rutba_forms.layout)

Positions and sizes are in CSS pixels from the form's top left corner.

| Key | Value |
|---|---|
| `name` | The form's name. |
| `table` | The table it shows, a record at a time. |
| `title` | Words at the top. |
| `grid` | Optional. The grid controls snap to: 8. |
| `width`, `height` | Optional. The form's size. |
| `controls` | Array of controls, one for each field shown. |
| `subforms` | Optional array of subforms. |

A control:

| Key | Value |
|---|---|
| `kind` | `text`, `longtext`, `number`, `currency`, `date`, `datetime`, `check` (a Yes/No), `list` (a field that points at another table's record) or `auto` (an AutoNumber, shown, not typed). An `integer` field is a `number` control. |
| `field` | The field of `table` it shows. |
| `label` | Words beside it. Its label stands 136 pixels to its left, 128 wide. |
| `x`, `y`, `w`, `h` | Where it is and how big. |
| `readOnly` | Optional, `true` for a control that may not be typed in. |
| `lookup` | For a `list`: `{"table": "<table>", "field": "<field>", "show": "<field>" or ["<field>", ...]}`, the records it lists, the field whose value it keeps, and the field or fields it shows. |

A subform shows the records of another table that point at this one:

| Key | Value |
|---|---|
| `table` | The other table. |
| `field` | Its field that points at this form's table. |
| `to` | The field of this form's table it points at. |
| `columns` | Array of the other table's fields to show. |
| `label` | Words above it. |
| `x`, `y`, `w`, `h` | Where it is and how big. |

### 4.3 A report's layout (rutba_reports.layout)

| Key | Value |
|---|---|
| `name` | The report's name. |
| `title` | Words at the top. |
| `source` | `{"table": "<table>"}` or `{"query": "<saved query>"}`. |
| `columns` | Array of `{"field", "label", "align"}`: a field of the table or a heading of the query, the words above it, and `left` or `right`. Optional in each: `type`, one of the field types, for a query's column; `decimals`, places to show. |
| `group` | A field to group the records by, or `null`. Groups are ordered by the words the field shows (a pointing field shows the words of the record it points at). |
| `sort` | A field to sort by within a group, or `null`. |
| `desc` | `true` to sort largest first. An empty value stays at the foot either way. |
| `totals` | Array of `{"field", "fn"}`, `fn` one of `sum`, `avg`, `count`, `min`, `max`: totals for each group and for the whole. Empty values are left out, not counted as 0. An optional `type` is ignored. |

### 4.4 A print template's layout (rutba_documents.layout)

A print template prints one document for each record of its table (an
invoice for each record of Invoices).

| Key | Value |
|---|---|
| `kind` | `"document"`. |
| `name` | The template's name. |
| `title` | Optional. Words to show for it. |
| `table` | The table it prints a document for each record of. |
| `paper` | Optional: `a4`, `letter`, `a5`, `receipt-80` or `receipt-58` (till rolls 80 and 58 mm wide). |
| `letterhead` | Optional: `{"table": "<table>", "fields": ["<field>", ...]}`, the first record of a table of the business's own details, at the top of every page; without `fields`, its text fields. |
| `where` | Optional expression: only the records it holds for are printed. |
| `blocks` | Array of blocks, from the top. |

Names in a template are the record's fields (`Date`), the same with the
table's name (`Invoices.Date`), a path through fields that point at other
records (`Customer.Address`, `Order.Customer.Name`; a pointing field named
alone shows the words of the record it points at), and the letterhead's
fields (`Address` or `Business Details.Address`). Expressions are those of
section 4.1, worked out with every name given as a value, so a template can
compute and do nothing else. An expression's answer "holds" unless it is
empty, 0, `"0"` or false.

Every block has a `type` and may have `if`, an expression: the block is
printed only when it holds.

| Type | Keys |
|---|---|
| `heading` | `text`: a line of large type, with `{Name}` blanks filled from the record. |
| `fields` | `columns`: 1 to 3, default 2. `items`: array of `{"field", "label", "type", "format", "if"}`, only `field` required; `format` is a `Format` pattern. |
| `lines` | The related records, as a table. `table`: the related table. `link`: its field that points at this record; or `on`: `{"<their field>": "<name on this record>"}`; with neither, the first relationship from it to this table. `columns`: array of `{"field"}`, or `{"expr", "as"}`, or `{"running": "<column>", "as", "start"}` (a running balance of another column, from `start`, an expression), each with optional `label`, `type`, `format`, `decimals`. Optional: `sort` (a field, `-` before it for largest first), `where` (expression over each line), `group` (a column to group under headings), `subtotals` (`true` by default, `false`, or an array of the columns to add up), `as` (a name: `<as> count` is the number of lines). |
| `totals` | `rows`: array of `{"label", "expr"}` with optional `if`, `type` (default `currency`), `format`, `strong` (bold), `decimals`. An expression may use the record's names, the sums of the lines' number columns under their names, and the labels of rows above it. |
| `text` | `text` with `{Name}` blanks; optional `small`. |
| `signatures` | `labels`: array of words, one line to sign on each; default `["Signed"]`. |
| `spacer` | `size`: room in pixels, default 16. |

### 4.5 A role's rights (rutba_roles.rights)

| Key | Value |
|---|---|
| `design` | `true`: the role may change the design and sees everything. Default `false`. |
| `tables` | `{"default": <level>, "each": {"<table>": <entry>}}`. `default` applies to tables not in `each`; when absent or not a level, `read`. |
| `hidden` | `{"queries": [...], "forms": [...], "reports": [...], "documents": [...]}`: names of saved items the role does not see. |

A level is one of `hidden`, `read`, `add`, `change`, `delete`, each allowing
what the ones before it allow. An entry of `each` is a level alone
(`"read"`), or `{"access": <level>, "own": <boolean>, "fields": {"<field>":
"hidden" | "read"}}`: `own` limits the role to the records its person added
(by `rutba_trail`), and `fields` hides fields or makes them read only. An
entry whose `access` is not a level is `hidden`. A saved item over a table
the role cannot see, or a report over a query it cannot see, is hidden too.

The role `Admin` is kept with the rights `{"design":true}`, and always has
every right whatever is kept. A database that is protected starts with two
more roles, which may be changed: `Staff` (`{"tables":{"default":"change"}}`)
and `Read only` (`{"tables":{"default":"read"}}`). Rutba Office writes the
other roles' rights with every key present.

## 5. Relationships

A relationship joins a field of one table (the "from" side, many records) to
a unique field of another (the "to" side, one record), usually its
AutoNumber. It is kept twice, and both must agree:

1. As a row of `rutba_relations`, which is what the design says.
2. As a `FOREIGN KEY` clause at the end of the "from" table's
   `CREATE TABLE`: `FOREIGN KEY ("<from_field>") REFERENCES "<to_tbl>"
   ("<to_field>") ON UPDATE CASCADE|RESTRICT ON DELETE CASCADE|RESTRICT`,
   by `cascade_update` and `cascade_delete`. This is what SQLite enforces.

Because SQLite cannot add a foreign key to a table that exists, adding or
removing a relationship makes the "from" table again: a new table is made
with the new clauses, the records are copied across with their rowids and
values, the old table is dropped and the new one renamed, and the
AutoNumber counter is carried over; foreign keys are checked when it is
done, and records that would point at nothing refuse the change. The same is
done to change a field's type, remove a field, or add a field that is
required or unique (a plain field is added with `ALTER TABLE ... ADD
COLUMN`). A writer may make these changes another way, as long as the
result is the same.

A relationship with a linked table (section 9) is a row of
`rutba_relations` with `cascade_update` and `cascade_delete` both 0 and no
`FOREIGN KEY` clause: queries join by it, and nothing enforces it, because
the records are the engine's.

## 6. The history of changes

Every change to the design is one SQLite transaction that also adds one row
to `rutba_steps`: the change and its step are written together or not at
all. Nothing is copied over anything, and records are changed in place.

| Label | Detail |
|---|---|
| `Create table T` | `{"table", "fields"}`, the fields as made. |
| `Add field F to T` | `{"table", "field"}` |
| `Change field F in T`, `Rename field F to G in T` | `{"table", "from", "to"}`, the field before and after; `"cleared": true` when values that did not fit the new type were let go empty. |
| `Remove field F from T` | `{"table", "field"}` |
| `Rename table T to U` | `{"from", "to"}` |
| `Delete table T` | `{"table"}` |
| `Relate T.F to U.G`, `Unrelate T.F from U.G` | `{"from", "to", "cascadeUpdate", "cascadeDelete"}` (`"linked": true` with a linked table); unrelating, `{"id"}`. |
| `Query Q kept`, `Query Q removed` | `{"query"}` |
| `Form F kept`, `Form F removed` | `{"form"}` |
| `Report R kept`, `Report R removed` | `{"report"}` |
| `Print template P kept`, `Print template P removed` | `{"document"}` |
| `Records of T named by F, G` | `{"table", "show"}` (the `shown` key of `rutba_meta`) |
| `Protect the database` | `{"admin"}`, the first Admin's name. |
| `Take the protection away` | `NULL` |
| `Role R kept`, `Role R removed` | `{"role"}` |
| `Person P added`, `Person P changed` | `{"person", "role"}` |
| `Person P removed` | `{"person"}` |
| `Link L kept`, `Link removed` | `{"link", "engine"}`; removing, `{"link"}` with the id. |
| `Link table T`, `Unlink table T` | `{"table", "link", "remote", "mode"}`; unlinking, `{"table"}`. |
| `Read T live`, `Keep a copy of T` | `{"table", "mode"}` |
| `Read T by F`, `Read all of T` | `{"table", "stamp"}` |

Not written as steps: records added, changed and deleted; a person's new
password; the design of a linked table as the engine describes it again; the
`pinned`, `folded`, `origin` and `origin.design` keys; and the bookkeeping of
`rutba_sync` and `rutba_clashes`.

How a reader treats the history: it is a record of what was done, for people
to read. The design is what the other design tables hold now, never what
replaying the steps would give, and a reader must not replay them. A reader
should show `label` as it is (it is English), may read `detail` for more,
and must not fail on a label or a detail it does not know. A writer that
changes the design should add a step in the same transaction, with the next
`n`, and must never change or remove a step. The steps carry no secret: a
step about a person names them and their role, never their password.

## 7. Design versions and upgrading

| Version | First written by | What it added |
|---|---|---|
| 1 | Rutba Office 2.12.0 | The file itself: `rutba_meta`, `rutba_tables`, `rutba_fields`, `rutba_relations`, `rutba_steps`; field types, relationships as foreign keys, the numbered history. |
| 2 | Rutba Office 2.12.0 | Saved queries and forms: `rutba_queries`, `rutba_forms`. |
| 3 | Rutba Office 2.12.0 | Reports: `rutba_reports`. |
| 4 | Rutba Office 2.13.0 | Print templates: `rutba_documents`. |
| 5 | Rutba Office 2.17.0 | People, roles and protection: `rutba_roles`, `rutba_people`, `rutba_trail` and its index, the `protected` key. |
| 6 | Rutba Office 2.17.0 | Tables linked live to a database engine: `rutba_links`, `rutba_linked`. Rutba Office 2.18.0 added, still as version 6, what version 7 marks (below), made only when first needed; so a version 6 file may hold any of it or none. |
| 7 | Rutba Office 2.19.0 | Tables kept as synced copies and tables read from other sources: `rutba_sync`, `rutba_clashes`, the `mode`, `stamp` and `synced` columns of `rutba_linked`, and the `file`, `rbase` and `json` kinds of link. In a version 7 file the two tables and the three columns are always there, empty when nothing is kept as a copy. A reader that knows only version 6 and meets a version 7 file must not change it (section 10, rule 2), so it cannot write to a copy as if it were read live. |

Every version keeps what the versions before it hold, unchanged. Nothing has
been renamed or removed, and no record is ever changed by an upgrade.

How an older file is upgraded: when Rutba Office opens a file to change it,
it runs every `CREATE TABLE IF NOT EXISTS` of the current version, which
makes the design tables the file lacks (empty), adds the columns of section
3 that `rutba_linked` lacks (`ALTER TABLE rutba_linked ADD COLUMN`, with the
definitions shown there), and then, when the file's `version` is lower,
writes the current version (7). That is the whole upgrade: from 6 to 7 it
adds `rutba_sync`, `rutba_clashes` and the three columns where a file does
not have them yet, and changes nothing that is there. A writer of another program must do the same: add the missing
tables and columns, then raise `version`; and it must never lower
`version`. A reader that only reads must not upgrade the file: it treats a
missing design table as empty.

## 8. Protected databases

### 8.1 What protection is, and what it is not

A protected database asks who you are as it opens, and gives each person
what their role allows (section 4.5). It is kept in the file as the
`protected` key of `rutba_meta` (`1`), the roles, the people with their
password hashes, and the trail.

Protection is not encryption. The records, the design and the trail are
stored in the file as they are in any other, and anyone who has the file can
read all of it, and change it, with any SQLite program. Protection is an
agreement that programs following this specification keep. To keep
something from someone, do not give them the file.

### 8.2 How passwords are kept

A password is at least 4 characters. It is never stored. `rutba_people.hash`
holds the text

```
scrypt$<N>$<r>$<p>$<salt>$<key>
```

six parts joined by `$`:

- `scrypt` names the function: scrypt (RFC 7914).
- `N`, `r`, `p` are its parameters in decimal. Rutba Office 2.19.0 uses
  N = 16384, r = 8, p = 1.
- `salt` is 16 random bytes from a cryptographically secure source, a new
  salt for every password set, in standard Base64 with padding (RFC 4648,
  section 4).
- `key` is the 32-byte key scrypt derives from the password (its UTF-8
  bytes) and the salt with those parameters, in the same Base64.

To check a password, decode `salt` and `key`, derive a key of the same
length as the decoded `key` with the `N`, `r` and `p` written in the text,
and compare the two in constant time. A text that does not have six parts
or does not start with `scrypt` matches no password. A reader must bound the
parameters before deriving (Node's default limit of 32 MiB of memory for
scrypt, which allows N = 16384 with r = 8, is a good bound), because the
file chooses them. A writer may use stronger parameters; it writes them in
the text.

### 8.3 What a program may do with a protected file

A program that follows this specification, when the database is protected:

- may read `rutba_meta` and the `name` column of `rutba_people`, to offer the
  names that may sign in, and nothing else of the people;
- must not show any record, field, query, form, report or print template
  until a person signs in with a name and a password that matches;
- after sign-in, must show and allow only what the person's role allows:
  tables, fields and records it can see; adding, changing and deleting by
  its level; saved items not hidden from it;
- must not change the design, and must not run a query written in SQL, for
  a role without `design`;
- must let only an Admin change roles and people, keep at least one Admin,
  and let anyone change their own password with their current one;
- must add a trail entry for each record added, changed or deleted;
- must not copy the database (Save As, export) for a role without `design`,
  as a copy would hold what the role cannot see;
- when an Admin takes the protection away, removes every row of
  `rutba_people`, keeps the roles and the trail, and sets `protected` to
  `0`.

Rutba Office keeps a role's limits beneath everything that reads: for the
signed-in person it makes, on its own connection, a TEMP VIEW for each
table with the same name, holding only the records and fields the role may
see. SQLite reads TEMP objects before the file's own, so queries, reports,
print templates and lookups read the views without knowing it, and writes go
to `main.<table>` after the role is checked. Nothing of this is in the file.

## 9. Linked tables, synced copies and other sources

A table can be linked instead of kept in the file: its records are a
database engine's (PostgreSQL, MySQL or MariaDB) or are read from another
source (a CSV file or a spreadsheet, another Data Files database, a web
address that answers JSON). The file keeps how to reach it, which part is
linked and the fields as last described.

**Secrets are never in the file.** An engine's password, a source
database's password and the values of web headers marked secret are kept by
the program on the computer, in that computer's protected store, never in
`rutba_links.settings` or anywhere else in the file. Rutba Office keeps them
in the profile's own keystore, whose key is wrapped by the operating system's
protection through Electron's `safeStorage` (DPAPI on Windows, the Keychain
on macOS, the secret service on Linux), under the names
`base-link:<link id>`, `base-source:<link id>` and
`base-source:<link id>:<header name>`. A file sent to someone else carries no
secret: they enter their own. A keystore that will not open on this computer
is put aside, and the database asks for its password again.

### 9.1 What each link keeps

`rutba_links.settings` and `rutba_linked.remote` by kind:

| Kind | settings | remote |
|---|---|---|
| `postgres`, `mysql` | `{"host", "port", "database", "user", "ssl"}`: text, a number or `null`, text, text, boolean. | `{"schema", "name"}`: the engine's schema (or database, for MySQL) and table. |
| `file` | `{"where", "sheet", "header", "every"}`: the file's path, the sheet's name or `null`, the row number of the headings (from 1). | `{"sheet"}` |
| `rbase` | `{"where", "person", "every"}`: the other database's path, and the name to sign in as there or `null`. | `{"item": {"kind": "table" \| "query", "name"}}` |
| `json` | `{"where", "headers", "records", "paging", "every"}`: the web address; headers `[{"name", "value"}]` or, for a secret one, `[{"name", "secret": true}]` with no value; the dotted path to the list of records in the answer (`data.items`), or empty; paging `{"kind": "none"}`, `{"kind": "next", "next": <path to the next page's address, or null for a Link header>}` or `{"kind": "page", "param", "start"}`. | `{}` |

`every` is how often a source is read again, in minutes: 5, 15 (the
default) or 60, or 0 for only when asked.

### 9.2 Live tables and copies

- **Live** (`mode` `live`, engines only). Nothing of the records is in the
  file. Rutba Office reads them into a TEMP table of the same name on its
  own connection, so the rest of the database reads them as any table, and
  writes each change to the engine at once.
- **Copy** (`mode` `copy`; every source). The records are a real table of
  the file with the linked table's name, so the database works with no
  engine at all. Its columns are the fields with the affinity of their type
  (`INTEGER` for `autonumber`, `integer` and `yesno`; `REAL` for `number`
  and `currency`; `TEXT` for the rest) and no `CHECK`, since the values are
  the engine's. A whole-number key is `INTEGER PRIMARY KEY`, so a record's
  rowid is its key; any other key is `UNIQUE`. A source's copy is read only
  here, and each read replaces its records.

When the engine or source describes a table differently from what
`rutba_fields` holds, the fields are replaced and a copy's table is made
again with the records of the fields it keeps.

### 9.3 Catching up

A copy catches up when the database opens, every 15 seconds while it is on
screen, and when asked. First, the changes made here (`state` not null) go
to the engine by key, in `rid` order. A record changed here whose engine
version changed too since it was last seen is a clash, and the engine wins:
its version is put here and the one from here is kept in `rutba_clashes`.
Then the engine's records are read and compared with the copy by key and
hash: new ones are added, changed ones changed, and ones gone from the
engine removed, leaving alone any record changed here and not yet sent. With
a `stamp`, only records changed since the largest stamp in the copy are read,
and the engine's list of keys says which are gone.

The hash of a record is the SHA-1, in lower-case hexadecimal, of the UTF-8
JSON text of an array holding each field's value in the order of the
fields, each made text first: `null` stays `null`; a whole number as its
decimal digits; any other number to 15 significant digits, written as
JavaScript's `String(Number(v.toPrecision(15)))` writes it; true and false
as `"1"` and `"0"`; anything else as its text. `JSON.stringify` writes the
array, with no spaces.

## 10. Rules for readers and writers

A program that reads or writes Data Files databases:

1. **Recognises a file by its content** (section 1), not by its extension.
   An SQLite database that is not a Data Files database is another
   program's: a program should open it read only, and must not add design
   tables to it, change its journal mode or write to it unless the person
   asks for that. Rutba Office shows such a file's tables as they are, with
   a design worked out from its schema and kept in memory only, and makes a
   Data Files copy of it when asked (section 11).
2. **Refuses a newer design version.** When `version` is greater than the
   highest it knows, it must not change the file in any way, and should not
   open it except to read, saying that the file was made by a newer program.
   It must not lower `version`.
3. **Ignores what it does not know, and keeps it.** A `rutba_*` table it
   does not know, a column of a design table it does not know, a key of
   `rutba_meta` it does not know and a key of a JSON layout it does not know
   are to be left alone: not shown as a person's data, and kept, unchanged,
   when it saves. It changes a design row by updating the columns it means
   to change (never by deleting the row and writing it again), and a layout
   by changing the keys it means to change.
4. **Does not invent.** It writes only the field types, block types, link
   kinds and values this text lists for the version it writes.
5. **Keeps the file's promises.** It runs `PRAGMA foreign_keys = ON` before
   writing, writes values in the forms of section 2.2, keeps `rutba_fields`
   and the tables' columns in step, keeps a relationship's row and its
   `FOREIGN KEY` together, never reuses an AutoNumber, and makes each change
   to the design one transaction with its step.
6. **Never trusts what a file says about trust.** A file is data from
   whoever made it. In particular:
   - `protected` being absent or `0`, or a role's rights, decide only what a
     program shows; they prove nothing about who may read the file
     (section 8.1).
   - A password is accepted only when it matches its hash; the hash's
     parameters are bounded before use (section 8.2).
   - A query written in SQL is run read-only, one statement, a `SELECT`; an
     expression, a criterion or a template is worked out only through the
     grammar of section 4.1, never pasted into SQL.
   - The paths, hosts and web addresses in `rutba_links` are what the file
     asks for, not permission. A program should ask the person before it
     reaches a host, a web address or a file that a database it did not make
     names, and must not send a secret it keeps to any address other than
     the one the secret was given for.
   - `origin` does not prove where a database came from.
7. **Keeps no secret in the file** (section 9).

## 11. Where Rutba Office departs from this text

Rutba Office 2.19.0 writes files that follow this text, and follows rules 1
to 3 of section 10 as a reader and writer of files made elsewhere:

- **Another program's SQLite database** opens read only. Its rowid tables
  are shown with their records; each column is given the field type its
  declared type names (by SQLite's rules of affinity, with `DATE`,
  `DATETIME`, `TIMESTAMP` and `BOOLEAN` read as dates, times and Yes/No)
  where every value agrees, and long text where they do not; a column that
  holds binary values is shown read only; a foreign key of one column
  between two tables shown is a relationship. A table `WITHOUT ROWID`, a
  virtual table, and a table whose name a person could not give are not
  shown, and the window names them. This design is kept in `TEMP` tables
  on Rutba Office's own connection: nothing is written to the file and its
  journal mode is left as it is. The person may then allow changes, and
  each record added, changed or deleted goes into the file as it is made,
  the design still not changed; or make a Data Files copy, a new `.rbase`
  file with the same tables, fields and records (an `INTEGER PRIMARY KEY`
  becomes the AutoNumber with its numbers, a foreign key a relationship
  where the field it points at is the AutoNumber or unique, and a column of
  binary values is left out, said so), the file it came from left as it was.
- **A newer design version is refused.** A file whose `version` is greater
  than 7, or is not a whole number, is not opened, and nothing in it is
  changed. The message says which version made it and that a newer Rutba
  Office opens it. Mail merge, which only reads, still reads such a file.
- **What is not known is kept.** A design change updates the rows of
  `rutba_fields` in place by name, so a column a newer version added keeps
  its values. Saving a query from the designer keeps every key of the query
  and of each column that the designer does not set. A role's rights and an
  engine link's settings keep the keys this text does not name.

Rutba Office 2.18.0 and earlier departed from rules 1 to 3: they opened any
SQLite database and added the design tables (empty) to it and switched it to
WAL mode; opened a file of a newer design version and kept its number; wrote
every row of `rutba_fields` for a table again on each design change, losing
the values of columns they did not know; and the query designer wrote each
column of a designed query again from the keys of section 4.1. Rutba Office
2.18.0 also wrote the features of version 7 as version 6 (section 7). Rutba
Office 2.17 reads a version 6 file's copies as if they were live, and does
not know the `file`, `rbase` and `json` kinds of link; neither it nor 2.18.0
refuses a version 7 file, so a version 7 file should not be opened with
them.

## 12. Sample files and the checker

`docs/data-files-format/samples/` holds four small databases made by
Rutba Office's own code (`tools/make-rbase-samples.mjs`, which makes them
again):

| File | What it shows |
|---|---|
| `plain-tables.rbase` | Two related tables with records: every field type, a required and a unique field, defaults, a relationship. |
| `query-form-report.rbase` | The same tables with a saved query (criteria, a calculation, totals), a form, a report grouped with totals, and a print template. |
| `copies-and-sources.rbase` | The same tables beside linked ones (section 9): a CSV file's rows kept as a copy, an engine's table kept as a copy with a change made here waiting to be sent, and a table read live, of which nothing but its design is in the file. The paths and hosts are examples. |
| `protected.rbase` | A protected database. The example sign-in is the name `Owner` with the password `example-password`, and the person `Clerk` (role `Staff`) has the password `example-clerk`. These passwords are examples, published so the file can be opened; never use them for anything else. |

The samples are in rollback journal mode (section 1). Their times and salts
differ each time they are made.

`tools/check-rbase.mjs` reads a file without changing it and says whether
it follows this text:

```
node tools/check-rbase.mjs <file.rbase>
```

It checks the SQLite file is sound, that the design version is one it knows,
that the design tables of that version and their columns are there, that
every field has an allowed type and matches its table's column, that every
relationship joins fields that are there and has its foreign key, that every
JSON column parses and has the keys this text requires, that the people and
roles of a protected database agree, and that no link keeps a secret. It
exits with 0 when the file follows this text, and with 1 and a list of the
problems when it does not (2 when it is used wrongly). It needs Node.js 22.5
or later, for `node:sqlite`, and nothing else.
