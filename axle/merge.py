import csv
import shutil

from .helpers import get_cached_path, get_tracked_sheets, validate_axle_project


def merge(verbose=False):
    """Update local copies of sheets based on cached copies.
    This does not read the XLSX spreadsheet."""
    # TODO: handle renamed sheets
    axle_dir = validate_axle_project()
    for sheet_title, details in get_tracked_sheets(axle_dir).items():
        local_path = details["Path"]
        cached_path = get_cached_path(axle_dir, sheet_title)
        if local_path.endswith(".csv"):
            with open(cached_path, "r") as fr:
                reader = csv.reader(fr, delimiter="\t")
                with open(local_path, "w") as fw:
                    writer = csv.writer(fw, delimiter=",", lineterminator="\n")
                    for row in reader:
                        writer.writerow(row)
        else:
            shutil.copy(cached_path, local_path)
