import csv
import os

from .exceptions import RmError
from .helpers import get_cached_path, get_tracked_sheets, set_logging, validate_axle_project


def rm(paths, keep=False, verbose=False):
    """Remove a set of sheets from the project.
    If keep_local=False, also delete the local tables."""
    set_logging(verbose)
    axle_dir = validate_axle_project()
    sheets = get_tracked_sheets(axle_dir)
    path_to_sheet = {
        os.path.abspath(details["Path"]): sheet_title for sheet_title, details in sheets.items()
    }

    # Check for either untracked or ignored sheets in provided paths
    untracked = []
    for p in paths:
        abspath = os.path.abspath(p)
        if abspath not in path_to_sheet:
            untracked.append(p)
    if untracked:
        raise RmError(f"unable to remove untracked file(s): {', '.join(untracked)}.")

    sheets_to_remove = {title: sheet for title, sheet in sheets.items() if sheet["Path"] in paths}
    # Make sure we are not deleting the last sheet
    if len(sheets) - len(sheets_to_remove) == 0:
        raise RmError(
            f"unable to remove {len(sheets_to_remove)} tracked sheet(s) - "
            "the spreadsheet must have at least one sheet."
        )

    # Maybe remove local copies
    if not keep:
        for p in paths:
            if os.path.exists(p):
                os.remove(p)

    # Remove the cached copies
    for sheet_title in sheets_to_remove.keys():
        cached_path = get_cached_path(axle_dir, sheet_title)
        if os.path.exists(cached_path):
            os.remove(cached_path)

    # Update sheet.tsv
    with open(f"{axle_dir}/sheet.tsv", "w") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=[
                "Title",
                "Path",
                "Frozen Rows",
                "Frozen Columns",
            ],
        )
        writer.writeheader()
        for title, sheet in sheets.items():
            if title not in sheets_to_remove.keys():
                sheet["Title"] = title
                writer.writerow(sheet)
