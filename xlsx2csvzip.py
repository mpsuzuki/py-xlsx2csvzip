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
import re
import json
import psutil
import time
from collections import Counter

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

# suffixes of the files to be zipped by default
SUFFIXES_IN_ZIP = [
  ".csv", ".png", ".emf", ".json"
]

SUFFIXES_OUT_ZIP = [
  "-pid.json"
]


def should_include_in_zip(args, workdir_path, target_path):
  if args.debug:
    return True

  workdir_path = Path(workdir_path)
  target_path  = Path(target_path)

  rel = target_path.relative_to(workdir_path)
  if rel.parts[0] == "scratch":
    return False

  if target_path.suffix not in SUFFIXES_IN_ZIP:
      return False

  for s in SUFFIXES_OUT_ZIP:
    if target_path.name.endswith(s):
      return False

  return True


# Check if the operating system is Windows
IS_WINDOWS = sys.platform.startswith("win32")


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
  parser.add_argument("xlsx", nargs="*", help="input xlsx file(s)")
  parser.add_argument("--files-from", help="read pathnames of XLSX from file")
  parser.add_argument(
    "--cwd",
    action="store_true",
    help="save auto-generated zip in the current directory instead of input directory",
  )
  parser.add_argument("--dir", help="output directory to emit all ZIP files")
  parser.add_argument("--rawdir", help="output directory to emit all raw CSV (for debugging)")
  parser.add_argument(
    "--xlsx2zip",
    action="store_true",
    help="auto-generate zip file named after the input xlsx",
  )
  parser.add_argument("--zip", help="output zip file path (explicit)")
  parser.add_argument(
    "--emf",
    action="store_true",
    help="export charts as high-resolution EMF files (requires clipboard)",
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
    "--no-clobber", "-nc",
    action="store_true",
    help="do not clobber existing file",
  )
  parser.add_argument("--fallback", action="store_true", help="try LibreOffice if Excel failed")
  parser.add_argument("--excel-timeout", type=int, default=60, help="timeout for Excel worker")
  parser.add_argument("--libre-timeout", type=int, default=10, help="timeout for LibreOffice worker")
  parser.add_argument("--debug", action="store_true", help="debug mode")
  parser.add_argument("--keep-work", action="store_true", help="keep working directory even if zip file is done")
  parser.add_argument("--bom", action="store_true",
    help="add BOM to CSV files",
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


def compose_output_pathname(workdir_path, basename, suffix):
  filename = f"{safe_filename(basename)}.{suffix}"
  return workdir_path / filename


def cleanup_dir(dir):
  dir_path = Path(dir)

  if dir_path.exists():
    shutil.rmtree(dir_path)

  dir_path.mkdir(parents=True, exist_ok=True)
  return dir_path


@contextmanager
def create_output_stream(args, dir_path, title, suffix):
  output_path = compose_output_pathname(dir_path, title, suffix)
  print(f"  create {output_path}", file=sys.stderr)
  if args.bom:
    csv_enc = "utf-8-sig"
  else:
    csv_enc = "utf-8"
  with open(output_path, "w", newline="", encoding=csv_enc) as f:
    yield f


def touch_stage(dir, stage):
  Path(dir, stage).touch()


def run_worker(args, worker_name, workdir_path, xlsx_path, int_timeout):
  worker_base = Path(worker_name).stem.removeprefix("xlsx2csv-")
  worker_json = workdir_path / (worker_base + "-pid.json")

  cmd = [ sys.executable,
    str(Path(__file__).resolve().parent / worker_name),
    "--dir", str(workdir_path),
    str(xlsx_path),
  ]
  try:
    touch_stage(workdir_path, f"{worker_name}-called-timeout{int_timeout}")
    print(f"call {worker_name}", file=sys.stderr,)
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=int_timeout,)

  except subprocess.TimeoutExpired as e:
    print(f"{worker_name} timed out", file=sys.stderr,)
    touch_stage(workdir_path, f"{worker_name}-expired-timeout")
    print(f"  lookup {worker_json}", file=sys.stderr,)
    if worker_json.exists():
      print(f"  found {worker_json}", file=sys.stderr,)
      touch_stage(workdir_path, f"{worker_base}-pidjson-found")
      worker_info = json.loads( worker_json.read_text(encoding="utf-8") )
      pid = worker_info["pid"]
      ps = psutil.Process(pid)
      if ps.name() == worker_info["name"] and abs(ps.create_time() - worker_info["create_time"]) < 1.0:
        print(f"  kill {pid}", file=sys.stderr,)
        touch_stage(workdir_path, f"{worker_base}-pid{pid}-kill")
        ps.kill()
    return False

  if result.returncode != 0:
    print(f"{worker_name} worker failed: {result.returncode}", file=sys.stderr,)
    print(f"  {result.stderr}", file=sys.stderr,)
    touch_stage(workdir_path, f"{worker_name}-failed")
    return False

  return True


def emit_properties_json(args, wb, workdir_path):
  props = wb.properties
  meta = {}
  for k, v in vars(props).items():
    if v is None:
      meta[k] = None
    elif hasattr(v, "isoformat"):
      meta[k] = v.isoformat()
    else:
      meta[k] = str(v)

  touch_stage(workdir_path, "emit-workbook-properties-json-try")
  Path(workdir_path, "workbook-properties.json").write_text(
    json.dumps(meta, indent=2, ensure_ascii=False, sort_keys=True,),
    encoding="utf-8",
  )
  touch_stage(workdir_path, "emit-workbook-properties-json-done")


def process_single_xlsx(args, xlsx_path_str):
  xlsx_path = Path(xlsx_path_str).resolve()
  print(f"Processing: {xlsx_path}", file=sys.stderr)

  summary = {}
  summary["xlsx"] = str(xlsx_path.name)
  summary["excel_timeout"] = args.excel_timeout
  summary["libre_timeout"] = args.libre_timeout
  timings = {}
  summary["timings"] = timings
  summary["status"] = "fail"
  summary["backend"] = "none"

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

  if zip_path is not None:
    workdir_path = zip_path.with_suffix(".work")
  elif args.rawdir is not None:
    workdir_path = Path(args.rawdir)
  else:
    return summary

  if args.no_clobber and zip_path.exists():
    print(f"destination {str(zip_path)} is existing, skip conversion", file=sys.stderr)
    return summary

  try:
    timings["load_workbook"] = time.monotonic()
    # openpyxl static side (Runs on any OS)
    wb = load_workbook(str(xlsx_path), data_only=False)
    timings["load_workbook"] = time.monotonic() - timings["load_workbook"]

    cleanup_dir(workdir_path)
    emit_properties_json(args, wb, workdir_path)

    timings["emit_format_formula"] = time.monotonic()
    for ws in wb.worksheets:
      with create_output_stream(args, workdir_path, ws.title, "format.csv") as o:
        write_csv(iter_format_rows(ws), o)

      with create_output_stream(args, workdir_path, ws.title, "formula.csv") as o:
        write_csv(iter_cell_rows(ws), o)
    timings["emit_format_formula"] = time.monotonic() - timings["emit_format_formula"]

    if args.cached:
      timings["emit_cached"] = time.monotonic()
      wb_cached = load_workbook(str(xlsx_path), data_only=True)
      for ws in wb_cached.worksheets:
        with create_output_stream(args, workdir_path, ws.title, "cached.csv") as o:
          write_csv(iter_cell_rows(ws), o)
      timings["emit_cached"] = time.monotonic() - timings["emit_cached"]

    excel_ok = False
    libre_ok = False
    timings["emit_value_excel"] = time.monotonic()
    excel_ok = run_worker(args, "xlsx2csv-excel.py", workdir_path, xlsx_path, args.excel_timeout)
    timings["emit_value_excel"] = time.monotonic() - timings["emit_value_excel"]

    if excel_ok:
      summary["status"] = "ok"
      summary["backend"] = "excel"
    elif args.fallback:
      timings["emit_value_libre"] = time.monotonic()
      libre_ok = run_worker(args, "xlsx2csv-libre.py", workdir_path, xlsx_path, args.libre_timeout)
      timings["emit_value_libre"] = time.monotonic() - timings["emit_value_libre"]
      if libre_ok:
        summary["status"] = "ok"
        summary["backend"] = "libre"

    Path(workdir_path, "summary.json").write_text(
      json.dumps(summary, indent=2, ensure_ascii=False, sort_keys=False,),
      encoding="utf-8",
    )

    if not excel_ok and not libre_ok:
      return summary

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED,) as zf:
      # print(f"  {zip_path} packing started", file=sys.stderr)
      touch_stage(workdir_path, "zip-packing-started")
      for path in Path(workdir_path).rglob("*"):
        rel_path = path.relative_to(workdir_path)
        # print(f"  test {path}", file=sys.stderr)
        if not should_include_in_zip(args, workdir_path, path):
          # print(f"  {zip_path.name} excludes {rel_path}", file=sys.stderr)
          continue
        print(f"  {zip_path.name} includes {rel_path}", file=sys.stderr)
        zf.write(path, rel_path,)
      touch_stage(workdir_path, "zip-packing-done")
      print(f"  {zip_path} completed", file=sys.stderr)

    if not args.keep_work:
      print(f"  remove {workdir_path}", file=sys.stderr)
      shutil.rmtree(workdir_path)

  except Exception as e:
    print("something goes wrong", file=sys.stderr)

  finally:
    print("finished", file=sys.stderr)

  return summary

def summarize(summaries):
  backend_counter = Counter()
  status_counter = Counter()
  timing_collection = {}
  for xlsx,summary in summaries.items():
    status_counter[summary["status"]] += 1
    backend_counter[summary["backend"]] += 1
    for k, v in summary["timings"].items():
      if not k in timing_collection:
        timing_collection[k] = [v]
      else:
        timing_collection[k].append(v)

  print(f"summary of {len(summaries)} results", file=sys.stderr)
  print("backend", file=sys.stderr)
  for k,v in backend_counter.items():
    print(f"  {k}={v}")
  print("timings", file=sys.stderr)
  for k,v in timing_collection.items():
    print(f"  timing {k}: avg={sum(v)/len(v):.2f} max={max(v):.2f}", file=sys.stderr)


def main():
  args = parse_args()
  summaries = {}

  if args.files_from:
    ctx = (
      nullcontext(sys.stdin)
      if args.files_from == "-"
      else open(args.files_from, "r", encoding="utf-8")
    )
    with ctx as fh:
      for xlsx_target in fh:
        xlsx_target = xlsx_target.rstrip("\r\n")
        print(f"open file {xlsx_target} specified by file list", file=sys.stderr)
        summary = process_single_xlsx(args, xlsx_target)
        summaries[xlsx_target] = summary
        summarize(summaries)
  else:
    print(args.xlsx)
    # Iterate over all provided XLSX targets
    for xlsx_target in args.xlsx:
      print(f"open file {xlsx_target} specified by argument", file=sys.stderr)
      summary = process_single_xlsx(args, xlsx_target)
      summaries[xlsx_target] = summary
      summarize(summaries)



if __name__ == "__main__":
  main()
