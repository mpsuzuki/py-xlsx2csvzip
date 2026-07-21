#!/usr/bin/env python

import argparse
import io
import sys
import math
import zipfile
import csv
from pathlib import Path
from itertools import zip_longest
from collections import Counter

import x2c_helper as x2c

DOT_SUFFIX_FORMAT = "." + x2c.SUFFIX_FORMAT
DOT_SUFFIX_FORMULA = "." + x2c.SUFFIX_FORMULA
DOT_SUFFIX_CACHED = "." + x2c.SUFFIX_CACHED
DOT_SUFFIX_VALUE = "." + x2c.SUFFIX_VALUE
DOT_SUFFIX_V_LIBRE = "." + x2c.value_suffix("libre")


def parse_args():
  parser = argparse.ArgumentParser(
    description="compare csv file in zip archive"
  )
  parser.add_argument("zips", nargs="*", help="input xlsx file(s)")
  parser.add_argument("--debug", action="store_true", help="debug mode")
  parser.add_argument("--limit", type=int,
                      help="limit the inspections to the first LIMIT files")
  parser.add_argument("--dir", type=str, help="directory of ZIP files")
  args = parser.parse_args()
  
  if args.dir:
    for zip_path in Path(args.dir).glob("*.zip"):
      args.zips.append(zip_path)

  return args


def looks_like_number(s):
  try:
    x = float(s)
    return math.isfinite(x)
  except ValueError:
    return False


def common_digits(f_a, f_b):
  return math.floor(- math.log10(
    abs(f_a - f_b) / max(abs(f_a), abs(f_b))
  ))


def compare_csv2(args, csv_a, csv_b, counter):
  num = 0
  for i, (row_a, row_b) in enumerate(zip_longest(csv_a, csv_b, fillvalue=[])):
    for j, (cell_a, cell_b) in enumerate(zip_longest(row_a, row_b)):
      num += 1
      s_a = "" if cell_a is None else cell_a.strip()
      s_b = "" if cell_b is None else cell_b.strip()
      if s_a == s_b:
        if args.debug:
          print(f"{cell_a}, {cell_b} -> str_equal")
        counter["str_equal"] += 1
        continue
      elif looks_like_number(s_a) is False and looks_like_number(s_b) is False:
        if args.debug:
          print(f"{cell_a}, {cell_b} -> str_diff")
        counter["str_diff"] += 1
        continue
      elif looks_like_number(s_a) is False or looks_like_number(s_b) is False:
        if args.debug:
          print(f"{cell_a}, {cell_b} -> str_vs_num")
        counter["str_vs_num"] += 1
        continue

      f_a = float(s_a)
      f_b = float(s_b)
      if (f_a - f_b) == 0:
        if args.debug:
          print(f"{cell_a}, {cell_b} -> num_equal")
        counter["num_equal"] += 1
        continue

      tag = f"num_{common_digits(f_a, f_b)}"
      if args.debug:
        print(f"{cell_a}, {cell_b} -> {tag}")
      counter[tag] += 1
  return num


def summarize(dic_counter):
  for tag in ("excel_vs_cache", "excel_vs_libre", "cache_vs_libre"):
    print(f"{tag}:")
    counter = dic_counter[tag]
    keys = counter.keys()
    # keys_str = sorted(filter(lambda s: s.startswith("str") , keys))
    keys_num = sorted(
                 filter(lambda s: s.startswith("num") and s != "num_equal",
                        keys),
                 key=lambda s: int(s.removeprefix("num_"))
    )
    # keys_sorted = keys_str
    # if "num_equal" in keys:
    #  keys_sorted.append("num_equal")
    keys_sorted = list(filter(lambda s: s in keys, ["str_equal", "str_diff", "num_equal"]))
    keys_sorted += keys_num
    for key in keys_sorted:
      print(f"  {key}: {counter[key]}")

    print("")


def test_single_zip(args, zip_path, dic_counter):
  with zipfile.ZipFile(zip_path) as zf:
    files = zf.namelist()
    worksheets = set()
    for nm in files:
      if nm.endswith(DOT_SUFFIX_VALUE):
        worksheets.add(nm.removesuffix(DOT_SUFFIX_VALUE))
      elif nm.endswith(DOT_SUFFIX_FORMAT):
        worksheets.add(nm.removesuffix(DOT_SUFFIX_FORMAT))
      elif nm.endswith(DOT_SUFFIX_CACHED):
        worksheets.add(nm.removesuffix(DOT_SUFFIX_CACHED))

    for ws in worksheets:
      nm_value_excel = ws + DOT_SUFFIX_VALUE
      nm_value_libre = ws + DOT_SUFFIX_V_LIBRE
      nm_cached = ws + DOT_SUFFIX_CACHED

      if nm_value_excel not in files:
        continue
      elif nm_value_libre not in files:
        continue
      
      nm2csv = {}
      for nm in [nm_value_excel, nm_value_libre, nm_cached]:
        if nm not in files:
          continue

        with zf.open(nm) as raw:
          with io.TextIOWrapper(raw, encoding="utf-8-sig", newline="") as fh:
            nm2csv[nm] = list(csv.reader(fh))

      if nm_value_excel in nm2csv and nm_cached in nm2csv:
        r = compare_csv2(args, nm2csv[nm_value_excel], nm2csv[nm_cached],
                         dic_counter["excel_vs_cache"])
        # print(f"compare excel vs cache: {r} entries")

      if nm_value_excel in nm2csv and nm_value_libre in nm2csv:
        r = compare_csv2(args, nm2csv[nm_value_excel], nm2csv[nm_value_libre],
                         dic_counter["excel_vs_libre"])
        # print(f"compare excel vs libre: {r} entries")

      if nm_cached in nm2csv and nm_value_libre in nm2csv:
        r = compare_csv2(args, nm2csv[nm_cached], nm2csv[nm_value_libre],
                         dic_counter["cache_vs_libre"])
        # print(f"compare cache vs libre: {r} entries")


def main():
  args = parse_args()
  dic_counter = {
    "excel_vs_cache": Counter(),
    "excel_vs_libre": Counter(),
    "cache_vs_libre": Counter(),
  }

  if args.limit:
    total = args.limit
  else:
    total = len(args.zips)

  for i, zip_path in enumerate(args.zips, start=1):
    if i > total:
      break
    print(f"\r{i}/{total} ({100*i/total:5.1f}%) {zip_path}",
          end="", file=sys.stderr, flush=True,)
 
    test_single_zip(args, zip_path, dic_counter)
  summarize(dic_counter)

if __name__ == "__main__":
  main()
