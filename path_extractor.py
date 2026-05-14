## “””
QVW Log Path Extractor

Reads variable names from column A of your existing Excel file,
looks up the resolved file path for each from the QVW .log file,
and writes it into column B (Resolved Path).

- Column A is never modified
- Paths are written without square brackets
- Rows with no match in the log are left blank in col B

Usage:
1. Set LOG_FILE to your .log file path
2. Set EXCEL_FILE to your Excel file path
3. Set SHEET_NAME if different from default
4. Run: python extract_qvw_paths.py
“””

import re
import os

# ── Configuration ──────────────────────────────────────────────────────────────

LOG_FILE   = r”TORG_ETL.QVW.log”   # Path to your .log file
EXCEL_FILE = r”list.xlsx”           # Path to your Excel file
SHEET_NAME = “QVD Paths”            # Sheet tab name (update if different)
START_ROW  = 2                      # First data row (row 1 is the header)

# ── Parse Log File ─────────────────────────────────────────────────────────────

def parse_log(log_path):
“””
Line-by-line parser.

```
Looks for this pattern in the log:
    Let [VarName] = Peek('FileLocation', ...)
    ...
    SET vFilePath = <actual path>

Returns dict: { "VarName": "resolved\\path" }
"""
path_map = {}
current_var = None

let_re  = re.compile(r"Let\s+\[([^\]]+)\]\s*=\s*Peek\(['\"]FileLocation['\"]", re.IGNORECASE)
path_re = re.compile(r"SET\s+vFilePath\s*=\s*(.+)", re.IGNORECASE)

with open(log_path, "r", encoding="utf-8", errors="replace") as f:
    for line in f:
        let_match = let_re.search(line)
        if let_match:
            current_var = let_match.group(1).strip()

        if current_var:
            path_match = path_re.search(line)
            if path_match:
                raw_path = path_match.group(1).strip()
                # Remove any surrounding square brackets e.g. [\\path\to\file]
                clean_path = raw_path.strip("[]")
                path_map[current_var] = clean_path
                current_var = None  # reset after pairing

return path_map
```

# ── Update Excel ───────────────────────────────────────────────────────────────

def update_excel(path_map, excel_path, sheet_name, start_row):
try:
import openpyxl
except ImportError:
import subprocess, sys
subprocess.check_call([sys.executable, “-m”, “pip”, “install”, “openpyxl”, “-q”])
import openpyxl

```
wb = openpyxl.load_workbook(excel_path)

if sheet_name not in wb.sheetnames:
    print(f"ERROR: Sheet '{sheet_name}' not found.")
    print(f"  Available sheets: {wb.sheetnames}")
    return

ws = wb[sheet_name]

matched   = 0
unmatched = []

for row_idx in range(start_row, ws.max_row + 1):
    cell_a = ws.cell(row=row_idx, column=1).value
    if not cell_a:
        continue

    var_name = str(cell_a).strip()

    if var_name in path_map:
        ws.cell(row=row_idx, column=2).value = path_map[var_name]
        matched += 1
    else:
        unmatched.append(var_name)

wb.save(excel_path)
print(f"\nDone. {matched} paths written to column B in '{sheet_name}'.")

if unmatched:
    print(f"\nNo path found in log for {len(unmatched)} variable(s):")
    for u in unmatched:
        print(f"  - {u}")
```

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
if not os.path.exists(LOG_FILE):
print(f”ERROR: Log file not found → {LOG_FILE}”)
return

```
if not os.path.exists(EXCEL_FILE):
    print(f"ERROR: Excel file not found → {EXCEL_FILE}")
    return

print(f"Parsing: {LOG_FILE}")
path_map = parse_log(LOG_FILE)

if not path_map:
    print("No variable→path mappings found. Check that your log contains:")
    print("  Let [VarName] = Peek('FileLocation', ...)")
    print("  SET vFilePath = <path>")
    return

print(f"Found {len(path_map)} variable→path mappings in log.")
update_excel(path_map, EXCEL_FILE, SHEET_NAME, START_ROW)
```

if **name** == “**main**”:
main()