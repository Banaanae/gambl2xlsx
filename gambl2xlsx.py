import xlsxwriter
import xml.etree.ElementTree as ET
import sys
from pathlib import Path
import re


if len(sys.argv) == 1:
    print("Usage: gambl2xlsx.py [-v] [-o name] [-g N] input")
    print()
    print("  -v       Data is vertical instead of horizontal")
    print()
    print("  -o name  Name of the resulting .xlsx file")
    print("           If omitted, will use the name from input")
    print("           Will fail if file already exists")
    # print()
    # print("  -nh      No headings, will not include data labels or headings")
    print()
    print("  -g N     Number of rows (or columns if -v) between each dataset")
    print()
    print("  input    Path to the .gambl file")
    exit(1)

# Set up args for parsing
argLen = len(sys.argv)
args = sys.argv[1:argLen - 1]
inputArg = sys.argv[argLen - 1]

# Defaults
vertical = False
outname = inputArg.removesuffix(".gambl")
noHeadings = False
gap = 0

for idx, arg in enumerate(args):
    match arg:
        case "-v":
            vertical = True
        case "-o":
            outname = args[idx + 1]
        case "-nh":
            noHeadings = True
        case "-g":
            gap = int(args[idx + 1])
        case "":
            pass # Invalid option or value for another option, don't warn due to latter

# Validate input
try:
    inputf = Path(inputArg).read_text()
except:
    print("Error: Could not find file " + inputArg)
    exit(2)

if Path(inputArg).suffix != ".gambl":
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

outfile = xlsxwriter.Workbook(outname + ".xlsx")
results = outfile.add_worksheet()

row = 0
offset = 0
for line in dataRows:
    for col, val in enumerate(line):
        try:
            val = float(val)
        except:
            pass
        finally:
            if not(vertical):
                results.write(row + offset, col, val) # write docs are row, col
            else:
                results.write(col, row + offset, val)
    row += 1
    offset = row // 2 * gap

outfile.close()
