import csv
import logging
import os

from .exceptions import RmError
from .helpers import get_cached_path, get_tracked_sheets, set_logging, validate_axle_project


def rm(titles, keep_local=False, verbose=False):
    """Remove a set of sheets from the project.
    If keep_local=False, also delete the local tables."""
    set_logging(verbose)
    axle_dir = validate_axle_project()
    sheets = get_tracked_sheets(axle_dir)

    if set(titles) == set(sheets.keys()):
        raise RmError("You cannot remove all sheets from a project!")

    for sheet_title in titles:
        if sheet_title not in sheets:
            logging.error(f"'{sheet_title}' is not a tracked sheet! This will be ignored.")
            continue
        details = sheets[sheet_title]
        local_path = details["Path"]
        cached_path = get_cached_path(axle_dir, sheet_title)
        os.remove(cached_path)
        if not keep_local:
            logging.info(f"Removing local copy of '{sheet_title}' at {local_path}")
            os.remove(local_path)

    # Get remaining sheets
    sheet_rows = []
    for sheet_title, details in sheets.items():
        if sheet_title not in titles:
            details["Title"] = sheet_title
            sheet_rows.append(details)

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
