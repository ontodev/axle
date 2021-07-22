import csv
import datetime
import logging
import os
import re

from openpyxl import load_workbook
from .helpers import (
    get_cached_path,
    get_config,
    get_tracked_sheets,
    a1_to_rowcol,
    set_logging,
    validate_axle_project,
)


def fetch(verbose=False):
    """Update cached copies of sheets based on the XLSX spreadsheet. Do not update local copies."""
    set_logging(verbose)
    axle_dir = validate_axle_project()
    config = get_config(axle_dir)
    wb = load_workbook(config["Spreadsheet Path"])
    tracked_sheets = get_tracked_sheets(axle_dir)

    # TODO: handle renames, handle formatting, notes, etc.

    new_sheets = []
    sheet_frozen = {}
    for sheet_title in wb.get_sheet_names():
        print(sheet_title)
        if sheet_title not in tracked_sheets:
            new_sheets.append(sheet_title)
        sheet = wb.get_sheet_by_name(sheet_title)
        frozen = sheet.freeze_panes
        if frozen:
            row, col = a1_to_rowcol(frozen)
            sheet_frozen[sheet_title] = {
                "row": row - 1,
                "col": col - 1,
            }
        else:
            sheet_frozen[sheet_title] = {"row": 0, "col": 0}

        # TODO: get formatting from cells
        rows = []
        for row in sheet.iter_rows():
            cells = []
            for cell in row:
                cells.append(cell.value)
            rows.append(cells)

        # Write to cached copy
        cached_path = get_cached_path(axle_dir, sheet_title)
        with open(cached_path, "w") as fw:
            writer = csv.writer(fw, delimiter="\t", lineterminator="\n")
            writer.writerows(rows)

    # Get updated sheet details
    sheet_rows = []
    for st, details in tracked_sheets.items():
        details["Title"] = st
        details["Frozen Rows"] = sheet_frozen[st]["row"]
        details["Frozen Columns"] = sheet_frozen[st]["col"]
        sheet_rows.append(details)
    for ns in new_sheets:
        filepath = get_new_path(ns, config["Directory"], config["File Format"])
        logging.info(f"Adding new sheet '{ns}' with local path {filepath}")
        sheet_rows.append(
            {
                "Title": ns,
                "Path": filepath,
                "Frozen Rows": sheet_frozen[ns]["row"],
                "Frozen Columns": sheet_frozen[ns]["col"],
            }
        )

    # Update sheet.tsv
    with open(f"{axle_dir}/sheet.tsv", "w") as fw:
        writer = csv.DictWriter(
            fw,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Title", "Path", "Frozen Rows", "Frozen Columns"],
        )
        writer.writeheader()
        writer.writerows(sheet_rows)


def get_new_path(sheet_title, directory, file_format):
    """Create a distinct local sheet path for a sheet."""
    basename = re.sub(r"[^A-Za-z0-9]+", "_", sheet_title.lower()).strip("_")
    if directory:
        basename = os.path.join(directory, basename)
    # Make sure the path is unique - the user can change this later
    if os.path.exists(basename + ".tsv"):
        # Append datetime if this path already exists
        td = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = f"{basename}_{td}.{file_format}"
    else:
        filepath = basename + "." + file_format
    return filepath
