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
import tempfile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

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


def classify(cell):
  if cell.value is None:
    return "EMPTY"
  if is_date_format(cell.number_format):
    return "DATETIME"
  if cell.number_format == "General":
    return "GENERAL"
  if isinstance(cell.value, (int, float)):
    return "NUMBER"
  return "OTHER"


def iter_format_rows(ws):
  for row in ws.iter_rows():
    yield [classify(cell) for cell in row]


def iter_cell_rows(ws):
  for row in ws.iter_rows():
    yield ["" if cell.value is None else cell.value for cell in row]


def iter_value_rows(com_ws):
  used = com_ws.UsedRange
  rows = used.Rows.Count
  cols = used.Columns.Count

  for r in range(1, rows + 1):
    values = []
    for c in range(1, cols + 1):
      value = com_ws.Cells(r, c).Value
      if value is None:
        value = ""
      values.append(value)
    yield values


def write_csv(rows, stream):
  writer = csv.writer(stream)
  for row in rows:
    writer.writerow(row)


def safe_filename(name):
  trans = str.maketrans(
    {
      "\\": "_",
      "/": "_",
      ":": "_",
      "*": "_",
      "?": "_",
      '"': "_",
      "<": "_",
      ">": "_",
      "|": "_",
    }
  )
  return name.translate(trans)


def compose_output_pathname(args, basename, suffix):
  """Switches the target stream based on user arguments."""
  filename = ".".join([safe_filename(basename), suffix])

  outdir = Path(args.dir)
  outfile = outdir / filename
  return outfile


@contextmanager
def open_output_stream(args, outfile):
  print(f"  create {outfile}", file=sys.stderr)
  with open(outfile, "w", newline="", encoding="utf-8") as f:
    yield f


def process_with_libreoffice(args):
  """Evaluates formulas via LibreOffice headless conversion, exports value.csv,

  and extracts embedded charts via an HTML conversion hack.
  """
  print("  evaluating formulas and extracting charts via LibreOffice", file=sys.stderr)
  libre_cmd = get_libreoffice_command()

  xlsx_path = Path(args.xlsx)

  # Enforce 'value.csv' if user specifies --force-value-name
  csv_suffix = "value.csv" if args.force_value_name else "value_libre.csv"

  workdir = Path(args.dir)
  scratchdir = workdir / "scratch"
  scratchdir.mkdir(parents=True, exist_ok=True)

  # --- STEP 1: Formula Evaluation (Convert to XLSX) ---
  cmd_xlsx = [
    libre_cmd, "--headless",
    "--convert-to", "xlsx",
    "--outdir", str(scratchdir),
    str(xlsx_path)
  ]
  try:
    subprocess.run(cmd_xlsx, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
  except Exception as e:
    print(f"  Error: Failed to execute LibreOffice XLSX conversion: {e}", file=sys.stderr)
    return

  calculated_xlsx = scratchdir / f"{xlsx_path.stem}.xlsx"
  if not calculated_xlsx.exists():
    print("  Error: LibreOffice output file was not found.", file=sys.stderr)
      return

  wb_libre = load_workbook(str(calculated_xlsx), data_only=True)
  for ws in wb_libre.worksheets:
    with open_output_stream(args, zfile, ws.title, csv_suffix) as o:
      write_csv(iter_cell_rows(ws), o)

  # --- STEP 2: Chart Extraction Hack (Convert to HTML) ---
  # LibreOffice outputs embedded shapes/charts as sequential 'img0.png', 'img1.png'...
  cmd_html = [
    libre_cmd, "--headless",
    "--convert-to", "html",
    "--outdir", str(scratchdir),
    str(xlsx_path)
  ]
  try:
    subprocess.run(cmd_html, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
  except Exception as e:
    print(f"  Warning: Failed to execute LibreOffice HTML conversion for charts: {e}", file=sys.stderr)
    return

  # Scan and process all extracted chart images chronologically
  extracted_images = sorted(tmpdir_path.glob("*.png"))

  for i, img_path in enumerate(extracted_images, start=1):
    chart_filename = f"libre_chart_{i}.png"
    dest_path = workdir / chart_filename
    shutil.copy2(img_path, dest_path)
    print(f"  create {dest_path}", file=sys.stderr)


def main():
  process_single_xlsx(parse_args())

if __name__ == "__main__":
  main()
