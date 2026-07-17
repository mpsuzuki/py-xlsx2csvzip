# py-xlsx2csvzip
Python script to convert XLSX file to CSVs in ZIP.

## Usage
```
usage: xlsx2csvzip.py [-h]
                      [--cwd] [--rawdir RAWDIR] [--dir DIR]
                      [--xlsx2zip] [--zip ZIP]
                      [--emf]
                      [--cached] [--eval]
                      [--libre] [--force-value-name]
                      xlsx [xlsx ...]

Export XLSX workbook contents into a ZIP file or directory

positional arguments:
  xlsx                input xlsx file(s)

options:
  -h, --help          show this help message and exit
  --cwd               save auto-generated zip in the current directory instead of input directory
  --dir DIR           output directory to emit all ZIP files
  --rawdir RAWDIR     output directory to emit all raw CSV (for debugging)
  --xlsx2zip          auto-generate zip file named after the input xlsx
  --zip ZIP           output zip file path (explicit)
  --emf               export charts as high-resolution EMF files (requires clipboard)
  --cached            write cached.csv using values stored in workbook
  --eval              evaluate formulas with Excel (Windows only) and write value.csv
  --libre             evaluate formulas with LibreOffice and write value_libre.csv
  --force-value-name  force LibreOffice output to be named 'value.csv' instead of 'value_libre.csv'
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

## Setup for Win32
### Setup Your PowerShell environment
1. Prepare setpath-python.ps1 from ```setpath-python.ps1.skel```,
   by rewriting ```PYHOME``` to the directory where python.exe resides.
2. Start PowerShell.
3. ```Set-ExecutionPolicy -Scope Process Bypass```
4. ```. .\setpath-python.ps1```
5. Check ```pyinfo``` alias command.

### Setup Your VENV
1. ```python -m venv xlsx2csvzip```,
   you can replace the venv name by your preferred name.
2. ```. .\xlsx2csvzip\Scripts\Activate.ps1```
3. ```python -m pip install --upgrade pip```
4. ```pip install openpyxl pywin32 olefile```

