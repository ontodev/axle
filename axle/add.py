import csv
import logging
import ntpath
import os.path

from .exceptions import AddError
from .helpers import get_tracked_sheets, set_logging, validate_axle_project


def add(path, title=None, freeze_row=0, freeze_column=0, verbose=False):
    """Add a table (TSV or CSV) to the AXLE project. This updates sheet.tsv.
    This does not add the sheet itself to the linked XLSX file."""
    set_logging(verbose)
    axle_dir = validate_axle_project()
    sheets = get_tracked_sheets(axle_dir)

    if not title:
        # Create the sheet title from file basename
        title = ntpath.basename(path).split(".")[0]
    if title in sheets:
        raise AddError(f"'{title}' sheet already exists in this project")

    # Check if provided path is a directory
    new_paths = []
    if os.path.isdir(path):
        if title:
            raise AddError("You cannot use the -t/--title option when adding a directory")
        for f in os.listdir(path):
            if f.endswith(".tsv") or f.endswith(".csv"):
                new_paths.append(os.path.join(path, f))
    else:
        if not path.endswith(".tsv") and not path.endswith(".csv"):
            raise AddError(f"File '{path}' must be a TSV or CSV table")
        new_paths.append(path)

    if not new_paths:
        raise AddError(f"No TSV or CSV tables exist in directory '{path}'")

    paths = {x["Path"]: t for t, x in sheets.items()}
    with open(f"{axle_dir}/sheet.tsv", "a") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Title", "Path", "Frozen Rows", "Frozen Columns"],
        )
        for p in new_paths:
            if p in paths:
                other_title = paths[path]
                raise AddError(f"Local table {path} already exists as '{other_title}'")

            if not title:
                cur_title = os.path.splitext(os.path.basename(p))[0]
                if cur_title in paths.keys():
                    raise AddError("Local table '{cur_title}' already exists")
            else:
                cur_title = title

            # Finally, add this TSV to sheet.tsv
            writer.writerow(
                {
                    "Title": cur_title,
                    "Path": p,
                    "Frozen Rows": freeze_row,
                    "Frozen Columns": freeze_column,
                }
            )

            logging.info(f"{cur_title} successfully added to project")
