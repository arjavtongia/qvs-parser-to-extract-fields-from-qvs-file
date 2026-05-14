import re
import csv
import sys
from pathlib import Path

def extract_join_blocks(qvs_text: str) -> list[dict]:
“””
Parses a QVS script and extracts metadata for every JOIN … LOAD … RESIDENT block.

```
Output columns:
    Tab Name        - value of the most recent $Tab {name} before this block
    Table Name      - table being joined INTO, e.g. ABC from INNER JOIN (ABC)
    QVD             - blank (reserved)
    Join            - join type: INNER / LEFT / RIGHT / OUTER
    Join Table Name - RESIDENT source table name
    Join QVD        - blank (reserved)
    Level           - blank (reserved)
"""
rows = []
current_tab = ""

# Tokenise line-by-line so we can track $Tab and JOIN blocks in order
lines = qvs_text.splitlines()
i = 0

# Patterns
tab_pattern      = re.compile(r'\$Tab\s+(.+)', re.IGNORECASE)
join_pattern     = re.compile(
    r'^(LEFT|RIGHT|OUTER|INNER)?\s*JOIN\s*\((\w+)\)', re.IGNORECASE
)
load_pattern     = re.compile(r'^\s*LOAD\b', re.IGNORECASE)
resident_pattern = re.compile(r'RESIDENT\s+(\w+)\s*;', re.IGNORECASE)

while i < len(lines):
    line = lines[i].strip()

    # Track current tab
    tab_match = tab_pattern.search(line)
    if tab_match:
        current_tab = tab_match.group(1).strip()
        i += 1
        continue

    # Detect a JOIN line
    join_match = join_pattern.match(line)
    if join_match:
        join_type  = (join_match.group(1) or "").upper() or "JOIN"
        table_name = join_match.group(2)

        # Scan forward for the RESIDENT clause (within the same block)
        resident_table = ""
        j = i + 1
        while j < len(lines):
            candidate = lines[j].strip()
            res_match = resident_pattern.search(candidate)
            if res_match:
                resident_table = res_match.group(1)
                break
            # Stop if we hit another JOIN or $Tab (new block started)
            if join_pattern.match(candidate) or tab_pattern.search(candidate):
                break
            j += 1

        rows.append({
            "Tab Name":        current_tab,
            "Table Name":      table_name,
            "QVD":             "",
            "Join":            join_type,
            "Join Table Name": resident_table,
            "Join QVD":        "",
            "Level":           "",
        })

        i = j  # jump past the block we just consumed
        continue

    i += 1

return rows
```

def write_csv(rows: list[dict], output_path: str):
fieldnames = [“Tab Name”, “Table Name”, “QVD”, “Join”, “Join Table Name”, “Join QVD”, “Level”]
with open(output_path, “w”, newline=””, encoding=“utf-8”) as f:
writer = csv.DictWriter(f, fieldnames=fieldnames)
writer.writeheader()
writer.writerows(rows)
print(f”Written {len(rows)} row(s) to {output_path}”)

def main():
if len(sys.argv) < 2:
print(“Usage: python qvs_join_extractor.py <path_to_qvs_file> [output.csv]”)
sys.exit(1)

```
qvs_path = sys.argv[1]
out_path = sys.argv[2] if len(sys.argv) > 2 else "qvs_join_output.csv"

qvs_text = Path(qvs_path).read_text(encoding="utf-8", errors="replace")
rows = extract_join_blocks(qvs_text)

if not rows:
    print("No JOIN blocks found.")
else:
    for r in rows:
        print(r)

write_csv(rows, out_path)
```

# ── Quick self-test ──────────────────────────────────────────────────────────

if **name** == “**main**”:
if len(sys.argv) == 1:          # no args → run the built-in test
SAMPLE_QVS = “””
$Tab Sales_Overview

SomeTable:
LOAD A, B, C
FROM someFile.csv;

INNER JOIN (ABC)
LOAD
A, B, C, D
RESIDENT XYZ;

LEFT JOIN (DEF)
LOAD
X, Y
RESIDENT PQR;

$Tab Finance_Tab

INNER JOIN (GHI)
LOAD
M, N
RESIDENT STU;
“””
rows = extract_join_blocks(SAMPLE_QVS)
print(f”{‘Tab Name’:<20} {‘Table Name’:<15} {‘Join’:<8} {‘Join Table Name’}”)
print(”-” * 65)
for r in rows:
print(f”{r[‘Tab Name’]:<20} {r[‘Table Name’]:<15} {r[‘Join’]:<8} {r[‘Join Table Name’]}”)
write_csv(rows, “/mnt/user-data/outputs/qvs_join_output_test.csv”)
else:
main()