import csv
import logging
import os

from .exceptions import ApplyError
from .helpers import (
    get_tracked_sheets,
    get_sheet_formats,
    get_sheet_notes,
    set_logging,
    update_formats,
    update_notes,
    validate_axle_project,
)

MESSAGE_HEADERS = ["table", "cell", "level", "rule id", "rule", "message", "suggestion"]


def apply(paths, verbose=False):
    set_logging(verbose)
    axle_dir = validate_axle_project()

    # TODO: support data validation tables
    message_tables = []
    for p in paths:
        if p.endswith("csv"):
            sep = ","
        else:
            sep = "\t"
        with open(p, "r") as f:
            # Get headers and rows
            reader = csv.DictReader(f, delimiter=sep)
            headers = [x.lower() for x in reader.fieldnames]
            rows = []
            for r in reader:
                rows.append({k.lower(): v for k, v in r.items()})

            for h in headers:
                if h not in MESSAGE_HEADERS:
                    raise ApplyError(f"The headers in table {p} are not valid for apply")
            message_tables.append(rows)

    apply_messages(axle_dir, message_tables)


def apply_messages(axle_dir, message_tables):
    """Apply one or more message tables (from dict reader) to the sheets as formats and notes."""
    tracked_sheets = get_tracked_sheets(axle_dir)
    sheet_to_formats = get_sheet_formats(axle_dir)

    # Remove any formats that are "applied" (format ID 0, 1, or 2)
    sheet_to_manual_formats = {}
    for sheet_title, cell_to_formats in sheet_to_formats.items():
        manual_formats = {}
        for cell, fmt in cell_to_formats.items():
            if int(fmt) > 2:
                manual_formats[cell] = fmt
            sheet_to_manual_formats[sheet_title] = manual_formats
    sheet_to_formats = sheet_to_manual_formats

    # Remove any notes that are "applied" (starts with ERROR, WARN, or INFO)
    sheet_to_notes = get_sheet_notes(axle_dir)
    sheet_to_manual_notes = {}
    for sheet_title, cell_to_notes in sheet_to_notes.items():
        manual_notes = {}
        for cell, note in cell_to_notes.items():
            if (
                not note.startswith("ERROR: ")
                and not note.startswith("WARN: ")
                and not note.startswith("INFO: ")
            ):
                manual_notes[cell] = note
        sheet_to_manual_notes[sheet_title] = manual_notes
    sheet_to_notes = sheet_to_manual_notes

    # Read the message table to get the formats & notes to add
    for message_table in message_tables:
        for row in message_table:
            # Check for cell location - skip if none
            cell = row.get("cell")
            if not cell or cell.strip() == "":
                continue
            cell = cell.upper()

            table = os.path.splitext(os.path.basename(row["table"]))[0]
            if table not in tracked_sheets:
                logging.warning(f"'{table}' is not a tracked sheet")
                continue

            if table in sheet_to_formats:
                cell_to_formats = sheet_to_formats[table]
            else:
                cell_to_formats = {}

            if table in sheet_to_notes:
                cell_to_notes = sheet_to_notes[table]
            else:
                cell_to_notes = {}

            # Check for current applied formats and/or notes
            current_fmt = -1
            current_note_author = None
            current_note_text = None
            if cell in cell_to_formats and int(cell_to_formats[cell]) <= 3:
                current_fmt = cell_to_formats[cell]
            if cell in cell_to_notes:
                current_note_text = cell_to_notes[cell].get("text")
                current_note_author = cell_to_notes[cell].get("author")
                if (
                    not current_note_text.startswith("ERROR")
                    and not current_note_text.startswith("WARN")
                    and not current_note_text.startswith("INFO")
                ):
                    # Not an applied note
                    current_note_text = None
                    current_note_author = None

            # Set formatting based on level of issue
            if "level" in row:
                level = row["level"].lower().strip()
            else:
                level = "error"
            if level == "error":
                cell_to_formats[cell] = 1
            elif level == "warn" or level == "warning":
                level = "warn"
                if current_fmt != 1:
                    cell_to_formats[cell] = 2
            elif level == "info":
                if current_fmt != 1 and current_fmt != 2:
                    cell_to_formats[cell] = 3

            message = None
            if "message" in row:
                message = row["message"]
                if message == "":
                    message = None

            suggest = None
            if "suggestion" in row:
                suggest = row["suggestion"]
                if suggest == "":
                    suggest = None

            rule_id = None
            if "rule id" in row:
                rule_id = row["rule id"]

            # Add the note
            rule_name = None
            if "rule" in row:
                rule_name = row["rule"]
                logging.info(f'Adding "{rule_name}" to {cell} as a(n) {level}')
            else:
                logging.info(f"Adding message to {cell} as a(n) {level}")

            # Format the note
            if rule_name:
                note = f"{level.upper()}: {rule_name}"
            else:
                note = level.upper()
            if message:
                note += f"\n{message}"
            if suggest:
                note += f'\nSuggested Fix: "{suggest}"'
            if rule_id:
                note += f"\nFor more details, see {rule_id}"

            # Add to dict
            if current_note_text:
                cell_to_notes[cell] = {"text": current_note_text, "author": current_note_author}
            else:
                cell_to_notes[cell] = {"text": note, "author": ""}

            sheet_to_formats[table] = cell_to_formats
            sheet_to_notes[table] = cell_to_notes

    # Update formats & notes TSVs
    update_notes(axle_dir, sheet_to_notes)
    update_formats(axle_dir, sheet_to_formats)
