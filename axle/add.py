import csv
import logging
import ntpath

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

    paths = {x["Path"]: t for t, x in sheets.items()}
    if path in paths:
        other_title = paths[path]
        raise AddError(f"Local table {path} already exists as '{other_title}'")

    # Finally, add this TSV to sheet.tsv
    with open(f"{axle_dir}/sheet.tsv", "a") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Title", "Path", "Frozen Rows", "Frozen Columns"],
        )
        # ID gets filled in when we add it to the Sheet
        writer.writerow(
            {
                "Title": title,
                "Path": path,
                "Frozen Rows": freeze_row,
                "Frozen Columns": freeze_column,
            }
        )

        logging.info(f"{title} successfully added to project")
