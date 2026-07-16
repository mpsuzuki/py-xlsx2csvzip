#!/usr/bin/env python3

import sys
import argparse
import shutil
import subprocess
import zipfile
from pathlib import Path
from contextlib import nullcontext


def parse_args():
  parser = argparse.ArgumentParser(
    description="Convert files to XLSX"
  )
  parser.add_argument("--files-from", help="read src/dst from file")
  return parser.parse_args()


IS_WINDOWS = sys.platform.startswith("win32")


def is_real_xlsx(path):
  try:
    with zipfile.ZipFile(path) as z:

      names = set(z.namelist())

      return (
        "[Content_Types].xml" in names
        and "xl/workbook.xml" in names
      )

  except Exception:
    return False


def excel_available():
  if not IS_WINDOWS:
    return False

  try:
    import win32com.client

    excel = win32com.client.DispatchEx(
      "Excel.Application"
    )
    excel.Quit()

    return True

  except Exception:
    return False


def find_libreoffice():
  candidates = []

  if IS_WINDOWS:
    candidates.append(
      r"C:\Program Files\LibreOffice\program\soffice.exe"
    )

    candidates.append(
      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    )

  for p in candidates:
    if Path(p).exists():
      return p

  p = shutil.which("soffice")
  if p:
    return p

  return None


def normalize_by_copy(src, dst):
  print(
    f"copy     {src} -> {dst}",
      file=sys.stderr,
  )
  shutil.copy2(src, dst)


def normalize_by_excel(src, dst):
  import win32com.client

  print(
    f"excel    {src} -> {dst}",
    file=sys.stderr,
  )

  excel = win32com.client.DispatchEx(
    "Excel.Application"
  )

  try:
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(src))

    try:
      #
      # 51 = xlsx
      #
      wb.SaveAs(str(dst), FileFormat=51,)

    finally:
      wb.Close(False)

  finally:
    excel.Quit()


def normalize_by_libreoffice(src, dst):
  soffice = find_libreoffice()

  if soffice is None:
    raise RuntimeError(
      "LibreOffice not found"
    )

  print(
    f"libre    {src} -> {dst}",
    file=sys.stderr,
  )

  outdir = Path(dst).parent

  subprocess.run(
    [
      soffice,
      "--headless",
      "--convert-to",
      "xlsx",
      "--outdir",
      str(outdir),
      str(src),
    ],
    check=True,
  )

  generated = (
    outdir
    / (Path(src).stem + ".xlsx")
  )

  if generated != dst:
    if dst.exists():
      dst.unlink()

    generated.replace(dst)


def process_pair(src, dst, use_excel):
  src = Path(src).resolve()
  dst = Path(dst).resolve()

  dst.parent.mkdir(
    parents=True,
    exist_ok=True,
  )

  if is_real_xlsx(src):
    normalize_by_copy(src, dst)

  elif use_excel:
    normalize_by_excel(src, dst)

  else:
    normalize_by_libreoffice(src, dst)


def main():
    args = parse_args()

    use_excel = excel_available()

    if use_excel:
        print("backend: Excel", file=sys.stderr)

    else:
        print("backend: LibreOffice", file=sys.stderr)

    failed = False


    ctx = (
      nullcontext(sys.stdin)
      if args.files_from is None or args.files_from == "-"
      else open(args.files_from, "r", encoding="utf-8")
    )
    with ctx as fh:
      for line in fh:
        line = line.rstrip("\r\n")
        if not line:
          continue
        if line.startswith("#"):
          continue
        if "\t" not in line:
          continue

        src, dst = line.split("\t", 1)

        try:
          process_pair(src, dst, use_excel)

        except Exception as e:
          failed = True

          print(f"ERROR: {src}: {e}", file=sys.stderr)
    sys.exit(1 if failed else 0)

if __name__ == "__main__":
    main()
