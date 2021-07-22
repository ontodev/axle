import logging
import os
import sys

from argparse import ArgumentParser
from .add import add
from .exceptions import AxleError
from .fetch import fetch
from .helpers import get_version
from .init import init
from .merge import merge
from .push import push
from .rm import rm

add_msg = "Add a table (TSV or CSV) to the project"
fetch_msg = "Update cached copies of tables with sheets from spreadsheet"
init_msg = "Init a new AXLE project"
merge_msg = "Update tracked tables with cached copies of sheets"
pull_msg = "Update tracked tables with sheets from spreadsheet"
push_msg = "Update spreadsheet with tracked table contents"
rm_msg = "Remove a table from the project"


def usage():
    return f"""axle [command] [options] <arguments>
commands:
  add      {add_msg}
  fetch    {fetch_msg}
  help     Print this message
  init     {init_msg}
  merge    {merge_msg}
  pull     {pull_msg}
  push     {push_msg}
  rm       {rm_msg}
  version  Print the AXLE version"""


def main():
    parser = ArgumentParser(usage=usage())
    global_parser = ArgumentParser(add_help=False)
    global_parser.add_argument("-v", "--verbose", help="Print logging", action="store_true")
    subparsers = parser.add_subparsers(dest="cmd")

    sp = subparsers.add_parser("help", parents=[global_parser])
    sp.set_defaults(func=run_help)
    sp = subparsers.add_parser("version", parents=[global_parser])
    sp.set_defaults(func=version)

    # ------------------------------- add -------------------------------
    sp = subparsers.add_parser(
        "add",
        parents=[global_parser],
        description=add_msg,
        usage="axle add PATH [-t TITLE -d DESCRIPTION]",
    )
    sp.add_argument("path", help="Path to TSV or CSV to add")
    sp.add_argument("-t", "--title", help="Optional title of the sheet")
    sp.add_argument("-r", "--freeze-row", help="Row number to freeze up to", default="0")
    sp.add_argument("-c", "--freeze-column", help="Column number to freeze up to", default="0")
    sp.set_defaults(func=run_add)

    # ------------------------------- fetch -------------------------------
    sp = subparsers.add_parser(
        "fetch", parents=[global_parser], description=pull_msg, usage="axle fetch"
    )
    sp.set_defaults(func=run_fetch)

    # ------------------------------- init -------------------------------
    sp = subparsers.add_parser(
        "init",
        parents=[global_parser],
        description=init_msg,
        usage="axle init TITLE [-p PATH -d DIRECTORY -f FILE_FORMAT]",
    )
    sp.add_argument("title", help="Title of the project")
    sp.add_argument("-p", "--path", help="Optional path for XLSX file")
    sp.add_argument("-d", "--directory", help="Directory containing tables to init project with")
    sp.add_argument(
        "-f", "--file-format", default="tsv", help="Default format for new tables (TSV or CSV)"
    )
    sp.set_defaults(func=run_init)

    # ------------------------------- merge -------------------------------
    sp = subparsers.add_parser(
        "merge", parents=[global_parser], description=pull_msg, usage="axle merge"
    )
    sp.set_defaults(func=run_merge)

    # ------------------------------- pull -------------------------------
    sp = subparsers.add_parser(
        "pull", parents=[global_parser], description=pull_msg, usage="axle pull"
    )
    sp.set_defaults(func=run_pull)

    # ------------------------------- push -------------------------------
    sp = subparsers.add_parser(
        "push", parents=[global_parser], description=push_msg, usage="axle push"
    )
    sp.set_defaults(func=run_push)

    # -------------------------------- rm --------------------------------
    sp = subparsers.add_parser(
        "rm",
        parents=[global_parser],
        description=rm_msg,
        usage="axle rm SHEET_TITLE [SHEET_TITLE ...]",
    )
    sp.add_argument("titles", help="Titles of sheets to remove from project", nargs="+")
    sp.add_argument(
        "-k", "--keep", help="If included, keep local copies of sheets", action="store_true"
    )
    sp.set_defaults(func=run_rm)

    args = parser.parse_args()
    if not hasattr(args, "func"):
        print(usage())
        print("ERROR: a command is required")
        sys.exit(1)
    args.func(args)


def run_add(args):
    """Wrapper for add function."""
    try:
        add(
            args.path,
            title=args.title,
            freeze_row=args.freeze_row,
            freeze_column=args.freeze_column,
            verbose=args.verbose,
        )
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def run_fetch(args):
    """Wrapper for fetch function."""
    try:
        fetch(args.verbose)
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def run_help(args):
    """Wrapper for help function."""
    print(usage())


def run_init(args):
    """Wrapper for init function."""
    try:
        success = init(
            args.title,
            filepath=args.path,
            directory=args.directory,
            file_format=args.file_format,
            verbose=args.verbose,
        )
        if not success:
            # Exit with error status without deleting AXLE directory
            sys.exit(1)
    except AxleError as e:
        # Exit with error status AND delete new AXLE directory
        logging.critical(str(e))
        if os.path.exists(".axle"):
            os.rmdir(".axle")
        sys.exit(1)


def run_merge(args):
    """Wrapper for merge function."""
    try:
        merge(verbose=args.verbose)
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def run_pull(args):
    """Wrapper for pull function."""
    try:
        fetch(verbose=args.verbose)
        merge(verbose=args.verbose)
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def run_push(args):
    """Wrapper for push function."""
    try:
        push(verbose=args.verbose)
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def run_rm(args):
    """Wrapper for push function."""
    try:
        rm(args.titles, keep_local=args.keep, verbose=args.verbose)
    except AxleError as e:
        logging.critical(str(e))
        sys.exit(1)


def version(args):
    """Print AXLE version information."""
    v = get_version()
    print(f"AXLE version {v}")


if __name__ == "__main__":
    main()
