#!/usr/bin/env python3

import sys
import argparse
from pathlib import Path

parser = argparse.ArgumentParser(
  description="Scan XLSX files under Moodle-structured tree"
)
parser.add_argument("srcs", nargs="*", help="source directories to be scanned")
parser.add_argument("--debug", action="store_true", help="print debug message")
parser.add_argument(
  "--suffixes",
  default="xlsx,xls",
  help="suffixes to be scanned (default: xlsx,xls)"
)
parser.add_argument("--dst", help="destination directory of normalized files")

parser.add_argument("--prefix", help="prefix of the normalized filename")
parser.add_argument(
  "--username-prefix",
  action="store_true",
  help="add username as prefix of the normalized filename"
)

parser.add_argument("--postfix", help="suffix of the normalized filename")
parser.add_argument(
  "--username-postfix",
  action="store_true",
  help="add username as postfix of the normalized filename"
)
parser.add_argument("--output", "-o",
  required=True,
  help="pathname to write of this command output")

args = parser.parse_args()

if len(args.srcs) == 0:
  args.srcs.append(".")
if args.dst is None:
  args.dst = "."
args.suffixes = args.suffixes.split(",")
args.dst = Path(args.dst)

if args.debug:
  print(f"Suffixes to be processed: {args.suffixes}", file=sys.stderr)

with open(args.output, "w", encoding="utf-8") as fh:
  for src in args.srcs:
    path_src = Path(src)
    for p in path_src.rglob("*"):
      suffix = p.suffix.lower().lstrip('.')
      if suffix not in args.suffixes:
        if args.debug:
          print(f"skip {p}", file=sys.stderr)
        continue

      d = p.parent.name
      if d.endswith("_assignsubmission_file_"):
        username = d.split("_")[0]
      else:
        username = None

      stem_out = p.stem
      if username and args.username_prefix:
        stem_out = username + "_" + stem_out
      if args.prefix:
        stem_out = args.prefix + stem_out
      if username and args.username_postfix:
        stem_out = stem_out + "_" + username
      if args.postfix:
        stem_out = stem_out + args.postfix

      path_out = args.dst / Path(stem_out + ".xlsx")

      print(f"{p}\t{path_out}", file = fh)
