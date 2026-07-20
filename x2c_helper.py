import sys
import csv
import shutil
from pathlib import Path
from contextlib import contextmanager

# -----------------------------------------------------------------
# shared strings

STATUS_SUCCESS = "success"
STATUS_FAILURE = "failure"
STATUS_TIMEOUT = "timeout"

def is_success(s):
  return (s == STATUS_SUCCESS)

def is_failure(s):
  return (s == STATUS_FAILURE)

def is_timeout(s):
  return (s == STATUS_TIMEOUT)

SUFFIX_FORMAT  = "format.csv"
SUFFIX_FORMULA = "formula.csv"
SUFFIX_CACHED  = "cached.csv"
SUFFIX_VALUE   = "value.csv"

def value_suffix(backend):
  if backend == "excel" or backend is None or len(backend) == 0:
    return SUFFIX_VALUE
  else:
    return f"value_{backend}.csv"

def emit_value(backend):
  return f"emit_value_{backend}"

# -----------------------------------------------------------------
# translator from None to ""

def iter_cell_rows(ws):
  for row in ws.iter_rows():
    yield ["" if cell.value is None else cell.value for cell in row]

# -----------------------------------------------------------------
# emit csv from 2D-array via csv.writer()

def write_csv(rows, stream):
  writer = csv.writer(stream)
  for row in rows:
    writer.writerow(row)

# -----------------------------------------------------------------
# sanitizer for filenames

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

# -----------------------------------------------------------------
# return Path(dir + basename + suffix)

def compose_output_pathname(dir, basename, suffix):
  """Switches the target stream based on user arguments."""
  filename = ".".join([safe_filename(basename), suffix])

  outdir = Path(dir)
  outfile = outdir / filename
  return outfile

# -----------------------------------------------------------------
# return stream to outfile, with appropriate encoding

@contextmanager
def open_output_stream(outfile, bom = False):
  print(f"  create {outfile}", file=sys.stderr)
  if bom:
    csv_enc = "utf-8-sig"
  else:
    csv_enc = "utf-8"
  with open(outfile, "w", newline="", encoding=csv_enc) as f:
    yield f

# -----------------------------------------------------------------
# create a directory and returns its Path object

def ensure_output_dir(dir, recreate=False):
  outdir = Path(dir)
  if outdir.exists():
    if recreate or not outdir.is_dir():
      # print(f"rmtree {dir}", file=sys.stderr)
      shutil.rmtree(dir)
  # print(f"mkdir {dir}", file=sys.stderr)
  outdir.mkdir(parents=True, exist_ok=True)
  return outdir

# -----------------------------------------------------------------
# touch a file under specified dir

def touch_stage(dir, stage):
  Path(dir, stage).touch()
