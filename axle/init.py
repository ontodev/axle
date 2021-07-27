import csv
import json
import logging
import os

from openpyxl import Workbook
from .add import add
from .exceptions import InitError
from .helpers import get_version, set_logging
from .push import push


DEFAULT_FORMATS = {
    "1": {
        "alignment": {},
        "border": {"outline": True},
        "fill": {"fgColor": {"rgb": "FFF6C7C4", "tint": 0.0}, "patternType": "solid"},
        "font": {
            "color": {"theme": 1},
            "family": 2.0,
            "name": "Calibri",
            "scheme": "minor",
            "sz": 11.0,
        },
        "number_format": "General",
    },
    "2": {
        "alignment": {},
        "border": {"outline": True},
        "fill": {"fgColor": {"rgb": "FFFFEFA1", "tint": 0.0}, "patternType": "solid"},
        "font": {
            "color": {"theme": 1},
            "family": 2.0,
            "name": "Calibri",
            "scheme": "minor",
            "sz": 11.0,
        },
        "number_format": "General",
    },
    "3": {
        "alignment": {},
        "border": {"outline": True},
        "fill": {"fgColor": {"rgb": "FFB2E7F5", "tint": 0.0}, "patternType": "solid"},
        "font": {
            "color": {"theme": 1},
            "family": 2.0,
            "name": "Calibri",
            "scheme": "minor",
            "sz": 11.0,
        },
        "number_format": "General",
    },
}


def init(title, filepath=None, directory=None, file_format="tsv", verbose=False):
    set_logging(verbose)
    cwd = os.getcwd()
    if os.path.exists(".axle"):
        # Do not raise AxleError, or else .axle/ will be deleted
        logging.critical(f"AXLE project already exists in {cwd}/.axle/")
        return False

    if file_format.lower() not in ["tsv", "csv"]:
        raise InitError("Unknown default file format: " + file_format)

    logging.info(f"initializing AXLE project '{title}' in {cwd}/.axle/")
    os.mkdir(".axle")
    if not filepath:
        filepath = title.replace(" ", "_") + ".xlsx"
    write_data(title, filepath, directory=directory, file_format=file_format)

    if os.path.exists(filepath):
        print("A spreadsheet already exists at " + filepath)
        print("Run `axle pull` to get current sheets")
    else:
        # Create new XLSX spreadsheet
        wb = Workbook()
        wb.save(filepath)

    # Add all from provided directory
    if directory:
        files = sorted(os.listdir(directory))
        for p in files:
            if not p.endswith(".csv") and not p.endswith(".tsv"):
                continue
            # Add this to sheet.tsv
            add(os.path.join(directory, p))
        # Push local sheets
        push(verbose=verbose)
    return True


def write_data(title, filepath, directory=None, file_format="tsv"):
    """Create AXLE data files in .axle directory: sheet.tsv."""
    # Create the "tracked" directory
    os.mkdir(".axle/tracked")

    # Store AXLE configuration
    with open(".axle/config.tsv", "w") as f:
        writer = csv.DictWriter(f, delimiter="\t", lineterminator="\n", fieldnames=["Key", "Value"])
        v = get_version()
        writer.writerow({"Key": "AXLE", "Value": "https://github.com/ontodev/axle"})
        writer.writerow({"Key": "AXLE Version", "Value": v})
        writer.writerow({"Key": "Title", "Value": title})
        writer.writerow({"Key": "Spreadsheet Path", "Value": filepath})
        writer.writerow({"Key": "Directory", "Value": directory})
        writer.writerow({"Key": "File Format", "Value": file_format.lower()})

    with open(f".axle/note.tsv", "w") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Sheet Title", "Cell", "Note", "Author"],
        )
        writer.writeheader()

    with open(".axle/formats.json", "w") as f:
        f.write(json.dumps(DEFAULT_FORMATS, sort_keys=True, indent=4))

    with open(f".axle/format.tsv", "w") as f:
        writer = csv.DictWriter(
            f, delimiter="\t", lineterminator="\n", fieldnames=["Sheet Title", "Cell", "Format ID"],
        )
        writer.writeheader()

    # TODO: data validation

    # sheet.tsv contains sheet (table/tab) details from the spreadsheet
    with open(".axle/sheet.tsv", "w") as f:
        writer = csv.DictWriter(
            f,
            delimiter="\t",
            lineterminator="\n",
            fieldnames=["Title", "Path", "Frozen Rows", "Frozen Columns"],
        )
        writer.writeheader()
