import logging

from .exceptions import ClearError
from .helpers import get_tracked_sheets, get_sheet_notes, get_sheet_formats, set_logging, update_notes, update_formats, validate_axle_project


def clear_formats(axle_dir, sheet_title):
    """Remove all formats from a sheet."""
    sheet_formats = get_sheet_formats(axle_dir)
    if sheet_title in sheet_formats:
        logging.info(f"removing all formats from '{sheet_title}'")
        del sheet_formats[sheet_title]
    update_formats(axle_dir, sheet_formats)


def clear_notes(axle_dir, sheet_title):
    """Remove all notes from a sheet."""
    sheet_notes = get_sheet_notes(axle_dir)
    if sheet_title in sheet_notes:
        logging.info(f"removing all notes from '{sheet_title}'")
        del sheet_notes[sheet_title]
    update_notes(axle_dir, sheet_notes)


def clear(keyword, on_sheets=None, verbose=False):
    """Remove formats and/or notes from one or more sheets."""
    set_logging(verbose)
    axle_dir = validate_axle_project()

    tracked_sheets = get_tracked_sheets(axle_dir)
    if not on_sheets:
        # If no sheet was supplied, clear from all
        on_sheets = list(tracked_sheets.keys())

    # Check if the user supplied any non-tracked sheets
    untracked = []
    for st in on_sheets:
        if st not in tracked_sheets.keys():
            untracked.append(st)
    if untracked:
        raise ClearError(
            f"The following sheet(s) are not part of this project: " + ", ".join(untracked)
        )

    # TODO: clear data validation once we've added support for it
    if keyword == "formats":
        for st in on_sheets:
            clear_formats(axle_dir, st)
    elif keyword == "notes":
        for st in on_sheets:
            clear_notes(axle_dir, st)
    elif keyword == "all":
        for st in on_sheets:
            clear_formats(axle_dir, st)
            clear_notes(axle_dir, st)
    else:
        raise ClearError("Unknown keyword: " + keyword)
