#!/usr/bin/env python3

import sys
import argparse
# import zipfile
import json
import hashlib

from pathlib import Path
from collections import Counter
from contextlib import nullcontext

import x2c_helper as x2c


def parse_args():
  parser = argparse.ArgumentParser(
    description="Collect summary.json from ZIP files"
  )
  parser.add_argument("json", help="input json file")
  parser.add_argument("--output", "-o", type=str, default="-", help="output file")
  parser.add_argument("--show-top", type=int, default=10, help="show longest times")
  parser.add_argument("--filenames", action="store_true", help="show filenames")
  parser.add_argument("--hashnames", action="store_true", help="show hash values of filenames")

  args = parser.parse_args()

  if args.output == "-":
    args.outctx = nullcontext(sys.stdout)
  else:
    print(f"output:{args.output}")
    args.outctx = open(args.output, "w", encoding="utf-8")

  return args


def print_percentile(args, backend, smrs):
  with args.outctx as fo:
    evb = x2c.emit_value(backend)
    smrs = sorted(smrs, key=lambda smr: smr["timings"][evb])
    total = len(smrs)

    for p in [50, 90, 95, 99, 100]:
      i = max(0, int(total * p / 100) - 1)
      smr = smrs[i]
      print(f"{backend} {p:03d}% #{i:03d} {smr["timings"][evb]:06.2f}", file=fo)

    for i in range(max(0, total - args.show_top), total):
      smr = smrs[i]
      print(f"{backend} #{i:03d} {smr["timings"][evb]:06.2f}", end="", file=fo)
      if args.filenames:
        print(f"\t{smr['xlsx']}\t{smr['_zipfile']}", file=fo)
      elif args.hashnames:
        hn = hashlib.md5(smr['xlsx'].encode("utf-8")).hexdigest()[:8]
        print(f"\t{hn}", file=fo)
      else:
        print("", file=fo)

    print("", file=fo)


def main():
  args = parse_args()
  summaries_all = json.loads( Path(args.json).read_text("utf-8") )
  summaries = {}
  for smr in summaries_all:
    if "backend" not in smr:
      continue

    if "backend_results" in smr:
      # new style
      for b_res in smr["backend_results"]:
        b = b_res["name"]
        evb = x2c.emit_value(b)

        if not x2c.is_success(b_res["status"]):
          continue

        if b not in summaries:
          summaries[b] = []

        summaries[b].append(smr)
        if evb not in smr["timings"]:
          smr["timings"][evb] = b_res["elapsed"]

    else:
      # old style
      b = smr["backend"]
      if b in summaries:
        summaries[b].append(smr)
      else:
        summaries[b] = [ smr ]

  for b, smrs in summaries.items():
    print_percentile(args, b, smrs)


if __name__ == "__main__":
  main()

