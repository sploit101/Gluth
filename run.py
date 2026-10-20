#!/usr/bin/env python3
"""One entry point for a local Gluth document run on macOS or Windows."""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"


def run(script: str, *args: str) -> None:
    print(f"\n>>> {script}", flush=True)
    # UTF-8 output for every stage: on Windows a piped or background stage
    # otherwise uses cp1252 and crashes when it prints a symbol or a file
    # name with special characters.
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    subprocess.run([sys.executable, str(HERE / script), *args],
                   cwd=HERE, check=True, env=env)


def main() -> int:
    ap = argparse.ArgumentParser(description="Extract, index, map and review station file content")
    ap.add_argument("--input", type=Path, default=HERE / "sources",
                    help="customer document folder inside this checkout")
    ap.add_argument("--template", type=Path,
                    help="PowerPoint station file template; needed for station mapping")
    ap.add_argument("--station", help="station name supplied by a human")
    ap.add_argument("--file", help="only map requirements from filenames containing this text")
    ap.add_argument("--ocr", action="store_true", help="OCR scanned pages before indexing")
    ap.add_argument("--embed", action="store_true", help="build optional local neural index")
    ap.add_argument("--write", action="store_true",
                    help="write a PowerPoint copy after reviewing out/zuordnung.xlsx")
    ap.add_argument("--include-review", action="store_true",
                    help="also write mappings flagged for human review (requires --write)")
    ap.add_argument("--freigabe", type=Path,
                    help="reviewed mapping spreadsheet for --write; must match this run")
    args = ap.parse_args()

    if args.write and not (args.template and args.station):
        ap.error("--write requires --template and --station")
    if args.include_review and not args.write:
        ap.error("--include-review requires --write")
    if args.freigabe and not args.write:
        ap.error("--freigabe requires --write")
    if args.template and not args.station:
        ap.error("--template requires a human-supplied --station")
    if args.template and not args.template.is_file():
        ap.error(f"template not found: {args.template}")

    try:
        run("extract.py", "--input", str(args.input))
        if args.ocr:
            run("ocr.py")
            run("extract.py", "--input", str(args.input))
        chunks = OUT / "chunks.jsonl"
        if not chunks.exists() or not chunks.stat().st_size:
            print("No searchable text was extracted. Check out/manifest.json; "
                  "use --ocr for scans.", file=sys.stderr)
            return 2
        run("index.py", *(["--embed", "--fast"] if args.embed else []))
        run("requirements.py")
        if args.template:
            run("zielfelder_bauen.py", "--vorlage", str(args.template))
            mapping_args = ["--station", args.station]
            if args.file:
                mapping_args += ["--datei", args.file]
            run("zuordnen.py", *mapping_args)
            if args.write:
                output = OUT / f"Stationsakte_{datetime.now():%Y%m%d_%H%M%S_%f}.pptx"
                writer_args = ["--vorlage", str(args.template),
                               "--station", args.station, "--aus", str(output)]
                if not args.include_review:
                    writer_args.append("--nur-sichere")
                if args.freigabe:
                    writer_args += ["--freigabe", str(args.freigabe)]
                run("akte_schreiben.py", *writer_args)
    except subprocess.CalledProcessError as exc:
        print(f"Stage failed: {exc.cmd[1]} (exit {exc.returncode})", file=sys.stderr)
        return exc.returncode or 1

    print(f"\nDone. Review outputs in {OUT}")
    if args.template and not args.write:
        print("Review zuordnung.xlsx, then rerun with --write to create the PowerPoint copy.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
