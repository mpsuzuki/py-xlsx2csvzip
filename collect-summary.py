#!/usr/bin/env python3

import sys
import argparse
import zipfile
import json

from pathlib import Path
from collections import Counter


def parse_args():
  parser = argparse.ArgumentParser(
    description="Collect summary.json from ZIP files"
  )
  # parser.add_argument("zip", nargs="*", help="input xlsx file(s)")
  parser.add_argument("--dir", help="input directory of ZIP files")
  parser.add_argument("--output", "-o", help="output JSON file")

  return parser.parse_args()

def main():
  args = parse_args()
  summaries = []

  zipfiles = list(Path(args.dir).glob("*.zip"))
  total = len(zipfiles)
  for i, zip_path in enumerate(zipfiles, start=1):
    print(f"\r{i}/{total} ({100*i/total:5.1f}%)", end="", file=sys.stderr, flush=True,)
    try:
      with zipfile.ZipFile(zip_path) as zf:
        with zf.open("summary.json") as f:
          summary = json.load(f)
          summary["_zipfile"] = str(zip_path.name)
          summaries.append(summary)
    except KeyError:
      print(f"{zip_path}: summary.json not found", file=sys.stderr)
    except Exception as e:
      print(f"{zip_path}: {e}", file=sys.stderr)

  summaries_json = json.dumps(summaries, indent=2, ensure_ascii=False,)
  if args.output:
    Path(args.output).write_text(summaries_json, encoding="utf-8")
  else:
    print(summaries_json)

if __name__ == "__main__":
  main()

