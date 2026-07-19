#!/usr/bin/env python3

import argparse
from contextlib import nullcontext, contextmanager
import csv
import io
import os
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

from x2c_helper import iter_cell_rows, write_csv
from x2c_helper import compose_output_pathname
from x2c_helper import open_output_stream
from x2c_helper import ensure_output_dir
from x2c_helper import touch_stage

# Check if the operating system is Windows
IS_WINDOWS = sys.platform.startswith("win32")
IS_MAC = sys.platform.startswith("darwin")


def parse_args():
  parser = argparse.ArgumentParser(
    description="Export XLSX workbook contents into a ZIP file or directory"
  )
  parser.add_argument("xlsx",
    help="input xlsx file")
  parser.add_argument("--dir", required=True,
    help="output directory to emit csv/png/emf files"
  )
  parser.add_argument(
    "--cached",
    action="store_true",
    help="write cached.csv using values stored in workbook",
  )
  parser.add_argument(
    "--eval",
    action="store_true",
    help="evaluate formulas with Excel (Windows only) and write value.csv",
  )
  parser.add_argument(
    "--force-value-name",
    action="store_true",
    help="force LibreOffice output to be named 'value.csv' instead of 'value_libre.csv'",
  )
  parser.add_argument("--bom", action="store_true",
    help="add BOM to CSV files",
  )

  return parser.parse_args()


def get_libreoffice_command():
  """Returns the platform-specific command to execute LibreOffice."""
  if IS_WINDOWS:
    standard_path = Path("C:/Program Files/LibreOffice/program/soffice.exe")
    if standard_path.exists():
      return str(standard_path)
    return "soffice"
  elif IS_MAC:
    mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_path.exists():
      return str(mac_path)
    return "soffice"
  else:
    # Linux / WSL default
    return "soffice"


def process_with_libreoffice(args):
  """Evaluates formulas via LibreOffice headless conversion, exports value.csv,

  and extracts embedded images via an HTML conversion hack.
  """
  print("  evaluating formulas and extracting images via LibreOffice", file=sys.stderr)
  libre_cmd = get_libreoffice_command()

  xlsx_path = Path(args.xlsx).resolve()
  ensure_output_dir(args.dir)

  # Enforce 'value.csv' if user specifies --force-value-name
  csv_suffix = "value.csv" if args.force_value_name else "value_libre.csv"

  workdir = Path(args.dir)
  scratchdir = workdir / "scratch"
  if scratchdir.exists():
    shutil.rmtree(scratchdir)
  scratchdir.mkdir(parents=True, exist_ok=True)

  # --- STEP 1: Formula Evaluation (Convert to XLSX) ---
  cmd_xlsx = [
    libre_cmd, "--headless",
    "--convert-to", "xlsx",
    "--outdir", str(scratchdir),
    str(xlsx_path)
  ]
  touch_stage(args.dir, "libre-calculate-started")
  result_calc = subprocess.run(cmd_xlsx, capture_output=True, text=True)
  if result_calc.returncode != 0:
    print(f"  Warning: LibreOffice failed to recalculate XLSX: {result_calc.returncode}", file=sys.stderr)
    print(f"  LibreOffice message: {result_calc.stderr}", file=sys.stderr)
    return False

  calculated_xlsx = scratchdir / f"{xlsx_path.stem}.xlsx"
  if not calculated_xlsx.exists():
    print("  Error: LibreOffice output file was not found.", file=sys.stderr)
    print(f"  LibreOffice message: {result_calc.stderr}", file=sys.stderr)
    return False

  touch_stage(args.dir, "libre-calculate-finished")
  try:
    wb_libre = load_workbook(str(calculated_xlsx), data_only=True)
  except Exception as e:
    print(f"  Error: Failed to parse calculated XLSX by openpyxl: {e}", file=sys.stderr)
    return False

  for ws in wb_libre.worksheets:
    output_pathname = compose_output_pathname(args.dir, ws.title, csv_suffix)
    with open_output_stream(output_pathname, bom=args.bom) as o:
      write_csv(iter_cell_rows(ws), o)

  # --- STEP 2: Chart Extraction Hack (Convert to HTML) ---
  # LibreOffice outputs embedded shapes/charts
  cmd_html = [
    libre_cmd, "--headless",
    "--convert-to", "html",
    "--outdir", str(scratchdir),
    str(xlsx_path)
  ]

  touch_stage(args.dir, "libre-html-started")
  result_html = subprocess.run(cmd_html, capture_output=True, text=True)
  if result_html.returncode != 0:
    print(f"  Warning: LibreOffice failed to convert to HTML: {result_html.returncode}", file=sys.stderr)
    print(f"  LibreOffice message: {result_html.stderr}", file=sys.stderr)
    return False

  touch_stage(args.dir, "libre-html-finished")

  # Scan and process all extracted images chronologically
  extracted_images = sorted(scratchdir.glob("*.png"))
  if len(extracted_images) == 0:
    touch_stage(args.dir, "libre-html-no-png")
  else:
    digits = len(str(len(extracted_images)))

    for i, img_path in enumerate(extracted_images, start=1):
      extracted_filename = f"libre_extracted_{i:0{digits}d}.png"
      dest_path = workdir / extracted_filename
      shutil.copy2(img_path, dest_path)
      print(f"  create {dest_path}", file=sys.stderr)

  return True


def main():
  ok = process_with_libreoffice(parse_args())
  sys.exit(0 if ok else 1)


if __name__ == "__main__":
  main()
