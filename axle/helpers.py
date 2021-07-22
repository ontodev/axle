import csv
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
