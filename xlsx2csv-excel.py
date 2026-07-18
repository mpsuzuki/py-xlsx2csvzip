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
import psutil
import json

# Dynamically import Windows-specific libraries only when running on Windows
import win32com.client
import ctypes
from ctypes import wintypes


def parse_args():
  parser = argparse.ArgumentParser(
    description="Export XLSX workbook contents into a ZIP file or directory"
  )
  parser.add_argument("xlsx",
    help="input xlsx file")
  parser.add_argument("--dir", required=True,
    help="output directory to emit csv/png/emf files"
  )
  parser.add_argument("--emf", action="store_true",
    help="export charts as high-resolution EMF files (requires clipboard)",
  )
  parser.add_argument("--interactive", action="store_true",
    help="show Excel application window and dialogs",
  )
  parser.add_argument("--bom", action="store_true",
    help="add BOM to CSV files",
  )

  return parser.parse_args()


GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
GetWindowThreadProcessId.argtypes = [
  wintypes.HWND,
  ctypes.POINTER(wintypes.DWORD),
]
GetWindowThreadProcessId.restype = wintypes.DWORD

OpenClipboard = ctypes.windll.user32.OpenClipboard
CloseClipboard = ctypes.windll.user32.CloseClipboard
GetClipboardData = ctypes.windll.user32.GetClipboardData
EmptyClipboard = ctypes.windll.user32.EmptyClipboard

CF_ENHMETAFILE = 14
GDI_ERROR = 0xFFFFFFFF

GetEnhMetaFileBits = ctypes.windll.gdi32.GetEnhMetaFileBits
GetEnhMetaFileBits.argtypes = [
  wintypes.HANDLE,
  wintypes.UINT,
  ctypes.c_void_p
]
GetEnhMetaFileBits.restype = wintypes.UINT

DeleteEnhMetaFile = ctypes.windll.gdi32.DeleteEnhMetaFile
DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
DeleteEnhMetaFile.restype = wintypes.BOOL


def hwnd_to_pid(hwnd):
  pid = wintypes.DWORD()
  GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
  return pid.value


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
  if args.bom:
    csv_enc = "utf-8-sig"
  else:
    csv_enc = "utf-8"
  with open(outfile, "w", newline="", encoding=csv_enc) as f:
    yield f


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


def export_charts(args, com_ws):
  """Finds all charts and exports them via Excel. PNG is default. EMF is optional."""

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
    png_path = Path(args.dir).resolve() / png_filename
    if png_path.exists():
      png_path.unlink()
    try:
      # (png_path.parent / (png_path.name + ".started")).touch()
      chart.Export(FilterName="PNG", Filename=str(png_path))
      # (png_path.parent / (png_path.name + ".done")).touch()
    except Exception as e:
      print(f"  Failed to export chart {i} as PNG in {com_ws.Name}: {e}", file=sys.stderr)

    # 2. Export as EMF (Clipboard fallback route, only if --emf is specified)
    if args.emf:
      emf_filename = f"{base_img_name}.emf"
      emf_path = Path(args.dir).resolve() / emf_filename
      if emf_path.exists():  
        emf_path.unlink()
      try:
        # (emf_path.parent / (emf_path.name + ".started")).touch()
        # Copy vector graphic data to clipboard
        chart.CopyPicture(Appearance=1, Format=2)

        # Grab raw binary data safely
        emf_data = get_emf_bytes_from_clipboard()

        if emf_data:
          with open(emf_path, "wb") as f:
            f.write(emf_data)
            # (emf_path.parent / (emf_path.name + ".done")).touch()
            print(f"  create {emf_path}", file=sys.stderr)

      except Exception as e:
        print(f"  Failed to export chart {i} as EMF in {com_ws.Name}: {e}", file=sys.stderr)
        if emf_path.exists():
          emf_path.unlink()


def ensure_output_dir(args):
  outdir = Path(args.dir)
  outdir.mkdir(parents=True, exist_ok=True)
  return outdir


def touch_stage(args, stage):
  Path(args.dir, stage).touch()


def process_single_xlsx(args):
  xlsx_path = Path(args.xlsx).resolve()
  print(f"\nProcessing: {xlsx_path}", file=sys.stderr)

  outdir = ensure_output_dir(args)

  excel = None
  try:
    print("  dispatch Excel", file=sys.stderr)
    excel = win32com.client.DispatchEx("Excel.Application")

    pid = hwnd_to_pid(excel.Hwnd)
    psinfo = psutil.Process(pid)
    pidinfo = {
      "pid": pid,
      "hwnd": excel.Hwnd,
      "exe": psinfo.exe(),
      "cmdline": psinfo.cmdline(),
      "name": psinfo.name(),
      "create_time": psinfo.create_time(),
    }
    pidjson = Path(args.dir) / "excel-pid.json"
    pidjson.write_text(json.dumps(pidinfo, indent=2), encoding="utf-8")

    touch_stage(args, "excel-started")
    excel.Visible = args.interactive
    excel.DisplayAlerts = args.interactive

    com_wb = excel.Workbooks.Open(str(xlsx_path))
    touch_stage(args, "excel-wb-started")
    touch_stage(args, "excel-wb-opened")
    excel.CalculateFullRebuild()
    touch_stage(args, "excel-wb-calculated")

    for com_ws in com_wb.Worksheets:
      output_pathname = compose_output_pathname(args, com_ws.Name, "value.csv")
      with open_output_stream(args, output_pathname) as o:
        # Path(args.dir, output_pathname.name + ".started").touch()
        write_csv(iter_value_rows(com_ws), o)
        # Path(args.dir, output_pathname.name + ".done").touch()

      export_charts(args, com_ws)
    com_wb.Close(False)

  finally:
    if excel is None:
      print("  could not start Excel", file=sys.stderr)
      touch_stage(args, "excel-not-started")
    else:
      print("  quit Excel", file=sys.stderr)
      excel.Quit()
      touch_stage(args, "excel-finished")


def main():
  process_single_xlsx(parse_args())


if __name__ == "__main__":
  main()
