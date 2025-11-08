import xlsxwriter
import xml.etree.ElementTree as ET
import sys
from pathlib import Path
import re


if len(sys.argv) == 1:
    print("Usage: gambl2xlsx.py input [output]")
    print("  Input:  Path to the .gambl file")
    print("  Output: Name of the resulting .xlsx file")
    exit(1)

# Validate input
try:
    inputf = Path(sys.argv[1]).read_text()
except:
    print("Error: Could not find file " + sys.argv[1])
    exit(2)

if Path(sys.argv[1]).suffix != ".gambl":
    print("Warning: File appears to not be a gambl file, trying anyway")

inputf = re.sub("^.*?<", "<", inputf) # Clean header
inputf = re.sub(".*$", "", inputf)    # Clean footer

try:
    root = ET.fromstring(inputf)
except Exception as err:
    print("Error: Could not parse XML, did you supply the correct file?")
    print(err)
    exit(3)

datasets = root.findall("./DataSet")

if len(datasets) == 0:
    print("No datasets were found, exiting")
    exit(0)

dataRows = []
for dataset in datasets:
    sets = dataset.findall("./DataColumn/ColumnCells")
    for _set in sets:
        lists = _set.text.split("\n\n\n")
        for _list in lists:
            line = _list.split("\n")
            dataRows.append(line)

if (len(sys.argv) > 2):
    outname = sys.argv[2]
else:
    outname = sys.argv[1].removesuffix(".gambl")
outfile = xlsxwriter.Workbook(outname + ".xlsx")
results = outfile.add_worksheet()

row = 0
for line in dataRows:
    for col, val in enumerate(line):
        try:
            val = float(val)
        except:
            val = val
        finally:
            results.write(row, col, val)
    row += 1

outfile.close()
