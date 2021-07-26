import csv
import json
import logging
import os
import pkg_resources
import re

from .exceptions import AxleError


def a1_to_rowcol(label):
    """Adapted from gspead.utils."""
    m = re.compile(r"([A-Za-z]+)([1-9]\d*)").match(label)
    if m:
        column_label = m.group(1).upper()
        row = int(m.group(2))

        col = 0
        for i, c in enumerate(reversed(column_label)):
            col += (ord(c) - 64) * (26 ** i)
    else:
        return None
    return row, col


def col_to_a1(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def get_cached_path(axle_dir, sheet_title):
    """Return the path to the cached version of a sheet based on its title."""
    filename = re.sub(r"[^A-Za-z0-9]+", "_", sheet_title.lower())
    return f"{axle_dir}/tracked/{filename}.tsv"


def get_config(axle_dir):
    """Get the configuration for this project as a dict."""
    config = {}
    with open(f"{axle_dir}/config.tsv", "r") as f:
        reader = csv.reader(f, delimiter="\t", lineterminator="\n")
        for row in reader:
            config[row[0]] = row[1]
    for r in ["Spreadsheet Path", "Title"]:
        if r not in config:
            raise AxleError(f"AXLE configuration does not contain key '{r}'")
    return config


def get_format_dict(axle_dir):
    """Get a dict of numerical format ID -> the format dict."""
    if (
        os.path.exists(f"{axle_dir}/formats.json")
        and not os.stat(f"{axle_dir}/formats.json").st_size == 0
    ):
        with open(f"{axle_dir}/formats.json", "r") as f:
            fmt_dict = json.loads(f.read())
            return {int(k): v for k, v in fmt_dict.items()}
    return {}


def get_sheet_formats(axle_dir):
    """Get a dict of sheet ID -> formatted cells."""
    sheet_to_formats = {}
    with open(f"{axle_dir}/format.tsv", "r") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for row in reader:
            sheet_title = row["Sheet Title"]
            cell = row["Cell"]
            fmt = int(row["Format ID"])
            if sheet_title in sheet_to_formats:
                cell_to_format = sheet_to_formats[sheet_title]
            else:
                cell_to_format = {}
            cell_to_format[cell] = fmt
            sheet_to_formats[sheet_title] = cell_to_format
    return sheet_to_formats


def get_sheet_notes(axle_dir):
    """Get a dict of sheet ID -> notes on cells."""
    sheet_to_notes = {}
    with open(f"{axle_dir}/note.tsv") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for row in reader:
            sheet_title = row["Sheet Title"]
            cell = row["Cell"]
            note = row["Note"]
            author = row["Author"]
            if sheet_title in sheet_to_notes:
                cell_to_note = sheet_to_notes[sheet_title]
            else:
                cell_to_note = {}
            cell_to_note[cell] = {"text": note, "author": author}
            sheet_to_notes[sheet_title] = cell_to_note
    return sheet_to_notes


def get_tracked_sheets(axle_dir):
    """Get the current tracked sheets in this project from sheet.tsv as a dict of sheet title ->
    details. They may or may not have corresponding cached/local sheets."""
    sheets = {}
    with open(f"{axle_dir}/sheet.tsv", "r") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for row in reader:
            title = row["Title"]
            if not title:
                continue
            del row["Title"]
            sheets[title] = row
    return sheets


def get_version():
    try:
        return pkg_resources.require("ontodev-axle")[0].version
    except pkg_resources.DistributionNotFound:
        return "developer-version"


def set_logging(verbose):
    """Set logging for AXLE based on -v/--verbose."""
    if verbose:
        logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    else:
        logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")


def update_formats(axle_dir, sheet_formats, overwrite=False):
    """Update format.tsv with current formatting from XLSX."""
    current_sheet_formats = {}
    if not overwrite:
        current_sheet_formats = get_sheet_formats(axle_dir)
    fmt_rows = []
    for sheet_title, formats in sheet_formats.items():
        current_sheet_formats[sheet_title] = formats
    for sheet_title, formats in current_sheet_formats.items():
        for cell, fmt in formats.items():
            fmt_rows.append({"Sheet Title": sheet_title, "Cell": cell, "Format ID": fmt})
    with open(f"{axle_dir}/format.tsv", "w") as f:
        writer = csv.DictWriter(
            f, delimiter="\t", lineterminator="\n", fieldnames=["Sheet Title", "Cell", "Format ID"],
        )
        writer.writeheader()
        writer.writerows(fmt_rows)


def update_notes(axle_dir, sheet_notes, overwrite=False):
    """Update note.tsv with current remote notes.
    Remove any lines with a Sheet ID in removed_ids."""
    current_sheet_notes = {}
    if not overwrite:
        current_sheet_notes = get_sheet_notes(axle_dir)
    note_rows = []
    for sheet_title, notes in sheet_notes.items():
        current_sheet_notes[sheet_title] = notes
    for sheet_title, notes in current_sheet_notes.items():
        for cell, note in notes.items():
            note_rows.append(
                {
                    "Sheet Title": sheet_title,
                    "Cell": cell,
                    "Note": note["text"],
                    "Author": note["author"],
                }
            )
    with open(f"{axle_dir}/note.tsv", "w") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Sheet Title", "Cell", "Note", "Author"],
        )
        writer.writeheader()
        writer.writerows(note_rows)


def validate_axle_project():
    """Validate that there is a valid AXLE project in this or the parents of this directory. If not,
    raise an error. Return the absolute path of the .axle directory."""
    cur_dir = os.getcwd()
    axle_dir = None
    while cur_dir != "/":
        if ".axle" in os.listdir(cur_dir):
            axle_dir = os.path.join(cur_dir, ".axle")
            break
        cur_dir = os.path.abspath(os.path.join(cur_dir, ".."))

    if not axle_dir:
        raise AxleError("An AXLE project has not been initialized in this or parent directories!")
    for r in ["sheet.tsv"]:  # TODO: format.tsv, note.tsv, validation.tsv
        if not os.path.exists(f"{axle_dir}/{r}") or os.stat(f"{axle_dir}/{r}").st_size == 0:
            raise AxleError(f"AXLE directory '{axle_dir}' is missing {r}")
    return axle_dir
