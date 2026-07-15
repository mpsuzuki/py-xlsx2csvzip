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
import zipfile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

# Check if the operating system is Windows
IS_WINDOWS = sys.platform.startswith("win32")
IS_MAC = sys.platform.startswith("darwin")

# Dynamically import Windows-specific libraries only when running on Windows
if IS_WINDOWS:
  import win32com.client
  import ctypes
  from ctypes import wintypes

  OpenClipboard = ctypes.windll.user32.OpenClipboard
  CloseClipboard = ctypes.windll.user32.CloseClipboard
  GetClipboardData = ctypes.windll.user32.GetClipboardData
  EmptyClipboard = ctypes.windll.user32.EmptyClipboard

  CF_ENHMETAFILE = 14
  GDI_ERROR = 0xFFFFFFFF

  GetEnhMetaFileBits = ctypes.windll.gdi32.GetEnhMetaFileBits
  GetEnhMetaFileBits.argtypes = [wintypes.HANDLE, wintypes.UINT, ctypes.c_void_p]
  GetEnhMetaFileBits.restype = wintypes.UINT

  DeleteEnhMetaFile = ctypes.windll.gdi32.DeleteEnhMetaFile
  DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
  DeleteEnhMetaFile.restype = wintypes.BOOL


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


def get_emf_bytes_from_clipboard():
  """Extracts raw EMF bytes from the Windows clipboard safely."""
  if not IS_WINDOWS:
    return None
  if not OpenClipboard(None):
    return None
  try:
    hemf = GetClipboardData(CF_ENHMETAFILE)
    if not hemf:
      return None

    # Get the required buffer size first
    size = GetEnhMetaFileBits(hemf, 0, None)
    if size == 0 or size == GDI_ERROR:
      return None

    # Allocate buffer and retrieve raw binary data
    buf = ctypes.create_string_buffer(size)
    if GetEnhMetaFileBits(hemf, size, buf) == 0:
      return None

    return buf.raw
  finally:
    # Clear and release clipboard ownership immediately
    EmptyClipboard()
    CloseClipboard()


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


def parse_args():
  parser = argparse.ArgumentParser(
    description="Export XLSX workbook contents into a ZIP file or directory"
  )
  parser.add_argument("xlsx", nargs="+", help="input xlsx file(s)")
  parser.add_argument("--rawdir", help="output directory to emit all raw CSV (for debugging)")
  parser.add_argument("--dir", help="output directory to emit all ZIP files")
  parser.add_argument("--zip", help="output zip file path (explicit)")
  parser.add_argument(
    "--xlsx2zip",
    action="store_true",
    help="auto-generate zip file named after the input xlsx",
  )
  parser.add_argument(
    "--cwd",
    action="store_true",
    help="save auto-generated zip in the current directory instead of input directory",
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
    "--libre",
    action="store_true",
    help="evaluate formulas with LibreOffice and write value_libre.csv",
  )
  parser.add_argument(
    "--force-value-name",
    action="store_true",
    help="force LibreOffice output to be named 'value.csv' instead of 'value_libre.csv'",
  )
  parser.add_argument(
    "--emf",
    action="store_true",
    help="export charts as high-resolution EMF files (requires clipboard)",
  )

  return parser.parse_args()


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


@contextmanager
def open_output_stream(args, zfile, basename, suffix):
  """Switches the target stream based on user arguments."""
  filename = ".".join([safe_filename(basename), suffix])

  if zfile:
    print(f"    {filename}: adding to zip", file=sys.stderr)
    stream = io.StringIO()
    try:
      yield stream
    finally:
      zfile.writestr(filename, stream.getvalue().encode("utf-8"))
  elif args.rawdir:
    outdir = Path(args.rawdir)
    outdir.mkdir(parents=True, exist_ok=True)
    outfile = outdir / filename
    print(f"  create {outfile}", file=sys.stderr)
    with open(outfile, "w", newline="", encoding="utf-8") as f:
      yield f
  else:
    print(f"  {basename} : {suffix.upper()} ===", file=sys.stderr)
    yield sys.stdout


def export_charts(args, zfile, com_ws):
  """Finds all charts and exports them via Excel. PNG is default. EMF is optional."""
  if not IS_WINDOWS:
    return

  chart_objects = com_ws.ChartObjects()
  if chart_objects.Count == 0:
    return

  ws_prefix = safe_filename(com_ws.Name)

  for i, chart_obj in enumerate(chart_objects, start=1):
    chart = chart_obj.Chart

    title_part = ""
    if chart.HasTitle:
      raw_title = chart.ChartTitle.Text
      if raw_title:
        title_part = f"_{safe_filename(raw_title)}"

    base_img_name = f"{ws_prefix}.chart_{i}{title_part}"

    # 1. Export as PNG (Standard direct export)
    png_filename = f"{base_img_name}.png"
    temp_png_path = Path(os.environ.get("TEMP", ".")).resolve() / png_filename
    try:
      chart.Export(FilterName="PNG", Filename=str(temp_png_path))
      if temp_png_path.exists():
        if zfile:
          print(f"    {png_filename}: adding to zip", file=sys.stderr)
          zfile.write(str(temp_png_path), png_filename)
          temp_png_path.unlink()
        elif args.dir:
          outdir = Path(args.dir)
          outdir.mkdir(parents=True, exist_ok=True)
          dest_path = outdir / png_filename
          print(f"  create {dest_path}", file=sys.stderr)
          if temp_png_path != dest_path:
            if dest_path.exists():
              dest_path.unlink()
            temp_png_path.rename(dest_path)
        else:
          print(f"  [Chart Exported (PNG)]: {png_filename}", file=sys.stderr)
          temp_png_path.unlink()
    except Exception as e:
      print(f"  Failed to export chart {i} as PNG in {com_ws.Name}: {e}", file=sys.stderr)

    # 2. Export as EMF (Clipboard fallback route, only if --emf is specified)
    if args.emf:
      emf_filename = f"{base_img_name}.emf"
      temp_emf_path = Path(os.environ.get("TEMP", ".")).resolve() / emf_filename
      try:
        # Copy vector graphic data to clipboard
        chart.CopyPicture(Appearance=1, Format=2)

        # Grab raw binary data safely
        emf_data = get_emf_bytes_from_clipboard()

        if emf_data:
          with open(temp_emf_path, "wb") as f:
            f.write(emf_data)

          if temp_emf_path.exists() and temp_emf_path.stat().st_size > 0:
            if zfile:
              print(f"    {emf_filename}: adding to zip", file=sys.stderr)
              zfile.write(str(temp_emf_path), emf_filename)
              temp_emf_path.unlink()
            elif args.dir:
              outdir = Path(args.dir)
              outdir.mkdir(parents=True, exist_ok=True)
              dest_path = outdir / emf_filename
              print(f"  create {dest_path}", file=sys.stderr)
              if temp_emf_path != dest_path:
                if dest_path.exists():
                  dest_path.unlink()
                temp_emf_path.rename(dest_path)
            else:
              print(f"  [Chart Exported (EMF)]: {emf_filename}", file=sys.stderr)
              temp_emf_path.unlink()
          else:
            print(f"  Failed to generate valid EMF file for chart {i} in {com_ws.Name}.", file=sys.stderr)
        else:
          print(f"  Failed to retrieve EMF data from clipboard for chart {i}.", file=sys.stderr)

      except Exception as e:
        print(f"  Failed to export chart {i} as EMF in {com_ws.Name}: {e}", file=sys.stderr)
        if temp_emf_path.exists():
          temp_emf_path.unlink()


def process_with_libreoffice(args, zfile, xlsx_path):
  """Evaluates formulas via LibreOffice headless conversion, exports value.csv,

  and extracts embedded charts via an HTML conversion hack.
  """
  print("  evaluating formulas and extracting charts via LibreOffice", file=sys.stderr)
  libre_cmd = get_libreoffice_command()

  # Enforce 'value.csv' if user specifies --force-value-name
  csv_suffix = "value.csv" if args.force_value_name else "value_libre.csv"

  with tempfile.TemporaryDirectory() as tmpdir:
    tmpdir_path = Path(tmpdir)

    # --- STEP 1: Formula Evaluation (Convert to XLSX) ---
    cmd_xlsx = [
      libre_cmd, "--headless",
      "--convert-to", "xlsx",
      "--outdir", str(tmpdir_path),
      str(xlsx_path)
    ]
    try:
      subprocess.run(cmd_xlsx, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
      print(f"  Error: Failed to execute LibreOffice XLSX conversion: {e}", file=sys.stderr)
      return

    calculated_xlsx = tmpdir_path / f"{xlsx_path.stem}.xlsx"
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
      "--outdir", str(tmpdir_path),
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

      if zfile:
        print(f"    {chart_filename}: adding to zip", file=sys.stderr)
        zfile.write(str(img_path), chart_filename)
      elif args.dir:
        outdir = Path(args.dir)
        outdir.mkdir(parents=True, exist_ok=True)
        dest_path = outdir / chart_filename
        print(f"  create {dest_path}", file=sys.stderr)
        shutil.copy2(img_path, dest_path)
      else:
        print(f"  [LibreOffice Chart Extracted]: {chart_filename}", file=sys.stderr)


def process_single_xlsx(args, xlsx_path_str):
  xlsx_path = Path(xlsx_path_str).resolve()
  print(f"\nProcessing: {xlsx_path}", file=sys.stderr)

  # Determine the target ZIP file path
  zip_path = None
  if args.zip:
    zip_path = Path(args.zip)
  elif args.xlsx2zip:
    if args.cwd:
      zip_path = Path.cwd() / f"{xlsx_path.stem}.zip"
    elif args.dir:
      zip_path = Path(args.dir) / f"{xlsx_path.stem}.zip"
    else:
      zip_path = xlsx_path.with_suffix(".zip")

  zfile = None
  if zip_path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    zfile = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)

  try:
    # openpyxl static side (Runs on any OS)
    wb = load_workbook(str(xlsx_path), data_only=False)

    for ws in wb.worksheets:
      with open_output_stream(args, zfile, ws.title, "format.csv") as o:
        write_csv(iter_format_rows(ws), o)

      with open_output_stream(args, zfile, ws.title, "formula.csv") as o:
        write_csv(iter_cell_rows(ws), o)

    if args.cached:
      wb_cached = load_workbook(str(xlsx_path), data_only=True)
      for ws in wb_cached.worksheets:
        with open_output_stream(args, zfile, ws.title, "cached.csv") as o:
          write_csv(iter_cell_rows(ws), o)

    # 1. Process Excel COM side
    run_excel = args.eval or (not args.cached and not args.libre and IS_WINDOWS)

    if run_excel:
      if not IS_WINDOWS:
        print("  Skipping Excel COM processing (Not running on Windows)", file=sys.stderr)
      else:
        print("  dispatch Excel", file=sys.stderr)
        excel = win32com.client.DispatchEx("Excel.Application")
        try:
          excel.Visible = False
          excel.DisplayAlerts = False

          com_wb = excel.Workbooks.Open(str(xlsx_path))
          excel.CalculateFullRebuild()

          for com_ws in com_wb.Worksheets:
            with open_output_stream(args, zfile, com_ws.Name, "value.csv") as o:
              write_csv(iter_value_rows(com_ws), o)

            export_charts(args, zfile, com_ws)

          com_wb.Close(False)
        finally:
          print("  quit Excel", file=sys.stderr)
          excel.Quit()

    # 2. Process LibreOffice side
    run_libre = args.libre or (not args.cached and not args.eval and not IS_WINDOWS)

    if run_libre:
      process_with_libreoffice(args, zfile, xlsx_path)

  finally:
    if zfile:
      zfile.close()
      print(f"  Created ZIP archive: {zip_path}", file=sys.stderr)


def main():
  args = parse_args()

  # Iterate over all provided XLSX targets
  for xlsx_target in args.xlsx:
    process_single_xlsx(args, xlsx_target)


if __name__ == "__main__":
  main()
