#!/usr/bin/env python3

import argparse
from contextlib import nullcontext, contextmanager
import csv
import io
import os
import sys
from pathlib import Path
import zipfile

import win32com.client
from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

# Windows API definitions for safely extracting raw EMF bytes from clipboard
import ctypes
from ctypes import wintypes

OpenClipboard = ctypes.windll.user32.OpenClipboard
CloseClipboard = ctypes.windll.user32.CloseClipboard
GetClipboardData = ctypes.windll.user32.GetClipboardData
EmptyClipboard = ctypes.windll.user32.EmptyClipboard

# Windows internal constants
CF_ENHMETAFILE = 14
GDI_ERROR = 0xFFFFFFFF

# GDI32 API definitions for checking size and duplicating EMF data
GetEnhMetaFileBits = ctypes.windll.gdi32.GetEnhMetaFileBits
GetEnhMetaFileBits.argtypes = [wintypes.HANDLE, wintypes.UINT, ctypes.c_void_p]
GetEnhMetaFileBits.restype = wintypes.UINT

DeleteEnhMetaFile = ctypes.windll.gdi32.DeleteEnhMetaFile
DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
DeleteEnhMetaFile.restype = wintypes.BOOL


def get_emf_bytes_from_clipboard():
  """Extracts raw EMF bytes from the Windows clipboard safely."""
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
  parser.add_argument("--dir", help="output directory (for debugging)")
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
    help="also evaluate formulas with Excel and write value.csv",
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
  elif args.dir:
    outdir = Path(args.dir)
    outdir.mkdir(parents=True, exist_ok=True)
    outfile = outdir / filename
    print(f"  create {outfile}", file=sys.stderr)
    with open(outfile, "w", newline="", encoding="utf-8") as f:
      yield f
  else:
    print(f"  {basename} : {suffix.upper()} ===", file=sys.stderr)
    yield sys.stdout


def export_charts(args, zfile, com_ws):
  """Finds all charts and exports them. PNG is default. EMF is optional via --emf."""
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
    else:
      zip_path = xlsx_path.with_suffix(".zip")

  zfile = None
  if zip_path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    zfile = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)

  try:
    # openpyxl side
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

    # Determine whether to proceed to Excel COM side
    proceed_to_eval = (not args.cached) or (args.cached and args.eval)

    if proceed_to_eval:
      # Excel COM side
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

          # Export charts from the current worksheet
          export_charts(args, zfile, com_ws)

        com_wb.Close(False)
      finally:
        print("  quit Excel", file=sys.stderr)
        excel.Quit()

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
