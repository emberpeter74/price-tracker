"""Command line: pricetracker scan | report."""
from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

from .excel_io import InputError
from .report import build_report
from .scan import run_scan

USAGE = """Price tracker - compares your prices with competitors' web prices.

Commands:
  pricetracker scan      Look up competitor prices for every item. Writes a dated
                         'Prices ...xlsx' file and 'Data issues.xlsx' to the output folder.
  pricetracker report    Merge all 'Prices ...xlsx' files into 'Price comparison.xlsx'
                         with comparison tables and price graphs.

Options (both commands):
  --input  <folder>      Folder with Items.xlsx, Competitors.xlsx, Product links.xlsx
                         (default: 'input' next to the program)
  --output <folder>      Folder for the results (default: 'output' next to the program)

See the user guide for details."""


def base_dir() -> Path:
    """Folder of the exe when built, else the current folder."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def _say(message: str) -> None:
    try:
        print(message)
    except UnicodeEncodeError:
        print(message.encode("ascii", "replace").decode("ascii"))


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(prog="pricetracker", add_help=False)
    parser.add_argument("command", nargs="?", choices=["scan", "report"])
    parser.add_argument("--input")
    parser.add_argument("--output")
    try:
        args = parser.parse_args(argv)
    except SystemExit:
        args = None
    if args is None or args.command is None:
        _say(USAGE)
        return 2

    input_dir = Path(args.input) if args.input else base_dir() / "input"
    output_dir = Path(args.output) if args.output else base_dir() / "output"
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        if args.command == "scan":
            run_scan(input_dir, output_dir)
        else:
            path = build_report(output_dir)
            _say(f"Report written: {path}")
            if "(new)" in path.name:
                _say("NOTE: 'Price comparison.xlsx' was open in Excel, so the report was saved under "
                     "this other name. Close Excel and run the report again.")
        return 0
    except InputError as exc:
        _say(f"Cannot continue: {exc}")
        return 1
    except KeyboardInterrupt:
        _say("Stopped.")
        return 130
    except Exception:
        log = output_dir / "error.log"
        try:
            with open(log, "a", encoding="utf-8") as f:
                f.write(f"\n--- {datetime.now():%Y-%m-%d %H:%M:%S} {' '.join(argv)}\n{traceback.format_exc()}")
        except OSError:
            pass
        _say(f"Something went wrong that the program did not expect. Details were written to {log}. "
             "Please send that file to the supplier.")
        return 1
