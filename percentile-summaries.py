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
  parser.add_argument("json", help="input json file")
  parser.add_argument("--show-top", type=int, default=10, help="show longest times")
  parser.add_argument("--filenames", action="store_true", help="show filenames")
  return parser.parse_args()


def print_percentile(args, backend, smrs):
  evb = "emit_value_" + backend
  smrs = sorted(smrs, key=lambda smr: smr["timings"][evb])
  total = len(smrs)
  for p in [50, 90, 95, 99, 100]:
    i = max(0, int(total * p / 100) - 1)
    smr = smrs[i]
    print(f"{backend} {p:03d}% #{i:03d} {smr["timings"][evb]:06.2f}")

  for i in range(max(0, total - args.show_top), total):
    smr = smrs[i]
    print(f"{backend} #{i:03d} {smr["timings"][evb]:06.2f}", end="")
    if args.filenames:
      print(f"\t{smr['xlsx']}\t{smr['_zipfile']}")
    else:
      print("")

  print("")


def main():
  args = parse_args()
  summaries_all = json.loads( Path(args.json).read_text("utf-8") )
  summaries = {}
  for smr in summaries_all:
    if "backend" not in smr:
      continue
    b = smr["backend"]
    if b in summaries:
      summaries[b].append(smr)
    else:
      summaries[b] = [ smr ]

  for b, smrs in summaries.items():
    print_percentile(args, b, smrs)


if __name__ == "__main__":
  main()

