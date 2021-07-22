# AXLE for XLSX

**WARNING** this is a work in progress.

AXLE takes a set of TSV files and allows you to edit them as XLSX spreadsheets.

For editing Google Sheets, try [COGS Operates Google Sheets](https://github.com/ontodev/cogs).

## Setup

AXLE is not yet distributed. To install this package, clone the GitHub repository and run the installation from the new directory:
```
$ pip install -r requirements.txt
$ pip install -e .
```

To see a list of all commands:
```
axle help
```

## Overview

Since AXLE is designed to synchronize a set of tables and an XLSX spreadsheet,
we try to follow the familiar `git` interface and workflow:

- [`axle init`](#init) creates an `.axle` directory to store configuration data and creates an XLSX spreadsheet for the project
- [`axle add foo.tsv`](#add) starts tracking the `foo.tsv` table as a sheet
- [`axle rm foo.tsv`](#rm) stops tracking the `foo.tsv` table as a sheet
- [`axle push`](#push) pushes changes to tables to the project spreadsheet
- [`axle fetch`](#fetch) fetches the data from the spreadsheet and stores it in `.axle/`
- [`axle merge`](#merge) overwrites tables with the data from the spreadsheet
- [`axle pull`](#pull) combines fetch and merge

There is no step corresponding to `git commit`.

We recommend running `axle push` after updating a tracked table to keep the XLSX sheets in sync. 
When updating the XLSX spreadsheet, we recommend running `cogs pull` to keep the tables in sync.

### Logging

To print info-level logging messages (error and critical level messages are always printed), run any command with the `-v`/`--verbose` flag:

```
axle [command & opts] -v
```

Otherwise, most commands succeed silently.

---

## Commands

### `add`

Running `add` will begin tracking a TSV or CSV table. The table details (path, name/title) get added to `.axle/sheet.tsv`.
This does not immediately change the XLSX spreadsheet -- use `axle push` to push all tracked tables to the project spreadsheet.

```
axle add PATH [-t TITLE]
```

The `-t`/`--title` is optional. If not provided, the title of the sheet will be the base filename (e.g. `tables/foo.tsv` will be named `foo`).

You can also specify a number of rows and or columns to freeze:

```
axle add PATH -r FREEZE_ROW -c FREEZE_COLUMN
```

If you specify `-r 2 -c 1`, then the first two rows and the first column will be frozen once the sheet is pushed to the XLSX spreadsheet.
If these options are not included, no rows or columns will be frozen.

### `fetch`

Running `fetch` will sync the `.axle/tracked/` directory with all spreadsheet changes.

```
axle fetch
```

This will write all sheets in the spreadsheet to that directory as `{sheet-title}.tsv` - this will overwrite the existing sheets in `.axle/tracked/`, but will not overwrite the versions specified by their path.
If a new sheet has been added to the XLSX spreadsheet, this sheet will be added to `.axle/sheet.tsv`. 
To sync the local version of sheets with the data in `.axle/`, run [`axle merge`](#merge).

### `init`

Running `init` creates an `.axle` directory containing configuration data. This also creates a new XLSX file with the project title, if one does not already exist.

```
cogs init TITLE
```

By default, the XLSX spreadsheet will be created as `{title}.xlsx`, where the title is lowercase with spaces replaced with underscores.
You can specify a different path for the spreadsheet with the `-p`/`--path` option, but this must always end with `.xlsx`:

```
cogs init TITLE -p PATH
```

If a spreadsheet already exists at the path, all existing sheets from that spreadsheet will automaticallyy be added to your project.

You can also specify a directory containing a set of tables (TSV/CSV) to start your project with using the `-d`/`--directory` option:

```
cogs init TITLE -d DIRECTORY
```

Any new sheets that are added to the spreadsheet will be given a default format of TSV when running `axle fetch` or `axle pull`.
If a directory has been specified, they will be saved to that directory. If you want to save new sheets as CSVs instead, just include `-f csv`/`--format csv`.

### `pull`

Running `pull` will sync tables with sheets in the XLSX spreadsheet.
This combines `axle fetch` and `axle merge` into one step.

```
cogs pull
```

Note that if you make changes to a table without running `axle push`, then run `axle pull`, the changes **will be overwritten**.

### `push`

Running `push` will sync the spreadsheet with your tables.
This includes creating new sheets for any added tables (`axle add`) and deleting sheets for any removed tables (`axle rm`).
Any changes to the tables are also pushed to the corresponding sheets.

```
cogs push
```

### `merge`

Running `merge` will sync tables with data in the `.axle` directory after running `axle fetch`.

```
cogs pull
```

Note that if you make changes to a table without running `axle push`, then run `axle fetch && axle merge`, the changes **will be overwritten**.

### `rm`

Running `rm` will stop tracking one or more tables.
This will delete all copies of the included paths, unless you include the optional `-k`/`--keep` flag.

```
axle rm PATH [PATH...]
```

This does not delete the sheet(s) from the spreadsheet - use `axle push` to push all local changes to the XLSX spreadsheet.
