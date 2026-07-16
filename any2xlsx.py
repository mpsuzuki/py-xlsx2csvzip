#!/usr/bin/env python3

import sys
import argparse
import shutil
import subprocess
import zipfile
import olefile
from pathlib import Path
from contextlib import nullcontext


# check path is a binary (OLE) file including "Workbook" or "Book"
# which is supposed to be an old Excel binary file.
def is_ole_xls(path):
  try:
    ole = olefile.OleFileIO(path)
    return (ole.exists("Workbook") or ole.exists("Book"))
  except Exception:
    return False


def is_html(path):
  with open(path, "rb") as fh:
    head = fh.read(4096).lower()
  return (
    b"<html" in head
    or b"<!doctype html" in head
    or b"<head" in head
  )


def classify_html(path):
  text = Path(path).read_text(encoding="utf-8", errors="ignore").lower()

  if "sharepoint.com" in text:
    return "HTML_SHAREPOINT"
  elif "onedrive.com" in text:
    return "HTML_ONEDRIVE"
  else:
    return "HTML_UNKNOWN"


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


def parse_args():
  parser = argparse.ArgumentParser(
    description="Convert files to XLSX"
  )
  parser.add_argument("--verbose", "-v", action="count", default=0,
                      help="increase verbosity")
  parser.add_argument("--files-from", help="read src/dst from file")
  parser.add_argument("--interactive", action="store_true",
                      help="display Excel window")
  parser.add_argument("--fallback", action="store_true",
                      help="if first convertor is failed, try fallback")
  parser.add_argument("--libreoffice-path", default=None,
                      help="specify the path to LibreOffice")
  args = parser.parse_args()
  if args.libreoffice_path is None:
    args.libreoffice_path = find_libreoffice()

  return args


def normalize_by_copy(src, dst, args):
  if args.verbose > 2:
    print(f"copy     {src} -> {dst}", file=sys.stderr)
  shutil.copy2(src, dst)


def normalize_by_excel(src, dst, args):
  import win32com.client

  if args.verbose > 1:
    print(f"excel    {src} -> {dst}", file=sys.stderr)

  excel = win32com.client.DispatchEx(
    "Excel.Application"
  )

  try:
    excel.Visible = args.interactive
    excel.DisplayAlerts = args.interactive

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


def normalize_by_libreoffice(src, dst, args):
  soffice = args.libreoffice_path
  if soffice is None:
    raise RuntimeError("LibreOffice not found")

  if args.verbose > 1:
    print(f"libre    {src} -> {dst}", file=sys.stderr)

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

  generated = ( outdir / (Path(src).stem + ".xlsx"))

  if generated != dst:
    if dst.exists():
      dst.unlink()

    generated.replace(dst)


def process_pair(src, dst, use_excel, args):
  src = Path(src).resolve()
  dst = Path(dst).resolve()

  dst.parent.mkdir(parents=True, exist_ok=True)

  if is_real_xlsx(src):
    normalize_by_copy(src, dst, args)

  elif use_excel and is_ole_xls(src):
    try:
      normalize_by_excel(src, dst, args)
    except Exception as err:
      if args.fallback and args.libreoffice_path is not None:
        print(f"Excel failed for {src}, try LibreOffice", file=sys.stderr)
        normalize_by_libreoffice(src, dst, args)
      else:
        print(f"ERROR: {src}: {err}", file=sys.stderr)

  else:
    try:
      normalize_by_libreoffice(src, dst, args)
    except Exception as err:
      if is_html(src):
        html_type = classify_html(src)
        print(f"ERROR: {src} is {html_type}")
      else:
        print(f"ERROR: {src} is unknown")


def main():
  args = parse_args()

  use_excel = excel_available()

  if use_excel:
    print("backend: Excel", file=sys.stderr)

  elif args.libreoffice_path is not None:
    print("backend: LibreOffice", file=sys.stderr)

  else:
    print("backend: none", file=sys.stderr)
    sys.exit(1)


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
        process_pair(src, dst, use_excel, args)

      except Exception as e:
        failed = True
        print(f"ERROR: {src}: {e}", file=sys.stderr)

  sys.exit(1 if failed else 0)

if __name__ == "__main__":
    main()
