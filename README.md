# gambl2xlsx

Converts vernier graphical analysis file to excel sheet

## Usage

Run `pip install xlsxwriter`

```
gambl2xlsx.py [-v] [-o name] [-nh] [-g N] input

  -v       Data is vertical instead of horizontal

  -o name  Name of the resulting .xlsx file
           If omitted, will use the name from input
           Will fail if file already exists

  -g N     Number of rows (or columns if -v) between each dataset

  input    Path to the .gambl file
```

## Note

Only tested on files created with Vernier Graphical Analysis, other formats from other software *may* work, but untested.
