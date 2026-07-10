# py-xlsx2csvzip
Python script to convert XLSX file to CSVs in ZIP.

## What is this?
A Python script to convert XLSX file to CSVs in ZIP. For example,
the resulted ```sample.zip``` file from ```sample.xlsx```
would include:
```
Sheet1.format.csv
Sheet1.formula.csv
Sheet1.cached.csv
Sheet1.value.csv (← if you have Excel)
Sheet2.format.csv
...
```

The formats, formulas and cached values can be extracted by ```openpyxl``` without using Excel,
but to extract really evaluated value, you need to have Excel and ```pywin32```.

## Dependency
* et_xmlfile (in my case, 2.0.0)
* openpyxl (in my case, 3.1.5)
* pywin32 (in my case, 312)

## Usage
```
usage: xlsx2csvzip.py [-h] [--dir DIR] [--zip ZIP] [--xlsx2zip] [--cwd]
                      [--cached] [--eval] [--emf]
                      xlsx [xlsx ...]

Export XLSX workbook contents into a ZIP file or directory

positional arguments:
  xlsx        input xlsx file(s)

options:
  -h, --help  show this help message and exit
  --dir DIR   output directory (for debugging)
  --zip ZIP   output zip file path (explicit)
  --xlsx2zip  auto-generate zip file named after the input xlsx
  --cwd       save auto-generated zip in the current directory instead of
              input directory
  --cached    write cached.csv using values stored in workbook
  --eval      also evaluate formulas with Excel and write value.csv
  --emf       export charts as high-resolution EMF files (requires clipboard)
```
