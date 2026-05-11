"""
QVD Lineage Mapper
==================
Flow:
  1. Read QVD names from "Lineage Structure" sheet → Column D (Final Dashboard QVD Name)
  2. Look up each QVD name in "QVD List" sheet → Column A (FROM) to match,
     get resolved network path from Column B (QVD Path)
  3. Scan QVW folder for the file that CREATES/STORES that QVD
  4. Extract all table names from the QVW load script
  5. Append results into "QVD List" sheet only → new columns: QVW File(s), Source Table Names

USAGE:
    pip install openpyxl
    python qvd_to_excel_mapper.py

Set the three paths in main() before running.
"""

import re
import openpyxl
from pathlib import Path


# ─────────────────────────────────────────────
# 1. QVW SCRIPT EXTRACTION
# ─────────────────────────────────────────────

def extract_script_from_qvw(qvw_path: Path) -> str:
    """
    Decode a QVW binary file and return its embedded load script.
    QVW stores scripts as UTF-16-LE; falls back to UTF-8.
    """
    try:
        raw = qvw_path.read_bytes()
        try:
            text_utf16 = raw.decode("utf-16-le", errors="ignore")
        except Exception:
            text_utf16 = ""
        try:
            text_utf8 = raw.decode("utf-8", errors="ignore")
        except Exception:
            text_utf8 = ""
        return text_utf16 if text_utf16.count("LOAD") > text_utf8.count("LOAD") else text_utf8
    except Exception as e:
        print(f"  [WARN] Could not read {qvw_path.name}: {e}")
        return ""


# ─────────────────────────────────────────────
# 2. TABLE NAME PARSING
# ─────────────────────────────────────────────

def parse_table_names(script_text: str) -> list:
    """
    Extract Qlik table names from a load script.
    Catches labelled LOADs and SQL FROM clauses.
    """
    table_names = set()

    # Pattern 1: TableName:\nLOAD or SELECT
    labelled = re.findall(
        r'^\s*\[?([A-Za-z0-9_ ]+?)\]?\s*:\s*\n\s*(?:LOAD|SELECT)',
        script_text, re.MULTILINE | re.IGNORECASE
    )
    table_names.update(t.strip() for t in labelled)

    # Pattern 2: FROM schema."TableName"
    sql_from = re.findall(
        r'\bFROM\b\s+(?:\w+\.)?["\[]?([A-Za-z0-9_]+)["\]]?',
        script_text, re.IGNORECASE
    )
    table_names.update(t.strip() for t in sql_from if len(t) > 2)

    # Remove noise
    noise = {"LOAD", "SELECT", "WHERE", "WITH", "UR", "FROM", "AND", "OR", "NOT", "AS", "BY"}
    table_names = {
        t for t in table_names
        if not t.startswith("$") and not t.startswith("v") and t not in noise
    }
    return sorted(table_names)


# ─────────────────────────────────────────────
# 3. QVW CREATOR DETECTION
# ─────────────────────────────────────────────

def script_creates_qvd(script_text: str, qvd_name: str) -> bool:
    """
    Returns True if the QVW script GENERATES the given QVD.
    Checks STORE ... INTO, CALL StoreQVD(...), and fallback name match.
    """
    base = qvd_name.replace(".qvd", "").strip()
    escaped = re.escape(base)

    # STORE ... INTO ... name.qvd
    if re.search(rf'\bSTORE\b[^;]*\bINTO\b[^;]*{escaped}[^;]*\.qvd', script_text, re.IGNORECASE):
        return True
    # CALL StoreQVD(...name...)
    if re.search(rf'\bCALL\b\s+StoreQVD[^;]*{escaped}', script_text, re.IGNORECASE):
        return True
    # Fallback: name anywhere + STORE exists
    if re.search(escaped, script_text, re.IGNORECASE):
        if re.search(r'\bSTORE\b', script_text, re.IGNORECASE):
            return True
    return False


# ─────────────────────────────────────────────
# 4. QVD LIST LOOKUP BUILDER
# ─────────────────────────────────────────────

def extract_stem(raw: str) -> str:
    """
    Strip variable prefixes and extract filename stem from a QVD path.
    e.g. "$(vQVDExtractPath)MTTR\iAlertTicketFinal.qvd(qvd)" → "iAlertTicketFinal"
    """
    s = re.sub(r'\$\([^)]+\)', '', raw)
    s = s.replace("(qvd)", "").replace("(QVD)", "")
    s = Path(s.replace("\\", "/")).stem.strip()
    return s


def build_qvd_lookup(ws_qvdlist) -> dict:
    """
    Build lookup from QVD List sheet:
      {qvd_stem_lower: (original_stem, resolved_network_path, row_num)}
    Column A = FROM (variable path), Column B = QVD Path (resolved path)
    """
    lookup = {}
    for row_num in range(2, ws_qvdlist.max_row + 1):
        col_a = ws_qvdlist.cell(row=row_num, column=1).value
        col_b = ws_qvdlist.cell(row=row_num, column=2).value
        if not col_a or not col_b:
            continue
        stem = extract_stem(str(col_a).strip())
        if stem:
            lookup[stem.lower()] = (stem, str(col_b).strip(), row_num)

    print(f"Built QVD lookup with {len(lookup)} entries from 'QVD List'")
    return lookup


# ─────────────────────────────────────────────
# 5. MAIN
# ─────────────────────────────────────────────

def main():
    # ── SET THESE THREE PATHS BEFORE RUNNING ─────────────────────────
    EXCEL_PATH  = r"C:\path\to\Stability_Dashboard.xlsx"
    QVW_FOLDER  = r"\\vmwas1388330\QVDOCS\QAPM\QVDATA\Model"
    OUTPUT_PATH = r"C:\path\to\Stability_Dashboard_updated.xlsx"
    # ─────────────────────────────────────────────────────────────────

    LINEAGE_SHEET = "Lineage Structure"
    QVDLIST_SHEET = "QVD List"

    print("=" * 60)
    print("QVD Lineage Mapper")
    print("=" * 60)

    # Load workbook
    wb = openpyxl.load_workbook(EXCEL_PATH)
    for sheet in [LINEAGE_SHEET, QVDLIST_SHEET]:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")

    ws_lineage = wb[LINEAGE_SHEET]
    ws_qvdlist = wb[QVDLIST_SHEET]

    # ── Step 1: Read QVD names from Lineage Structure Column D ─────────
    print(f"\nStep 1: Reading from '{LINEAGE_SHEET}'...")
    TARGET_HEADER = "Final Dashboard QVD Name"
    qvd_col = None

    header_row = list(ws_lineage.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    for col_idx, cell_val in enumerate(header_row, start=1):
        if cell_val and str(cell_val).strip().lower() == TARGET_HEADER.lower():
            qvd_col = col_idx
            break

    if qvd_col is None:
        found = [str(v) for v in header_row if v]
        raise ValueError(
            f"Column '{TARGET_HEADER}' not found in row 2 of '{LINEAGE_SHEET}'.\n"
            f"Row 2 has: {found}"
        )

    print(f"         '{TARGET_HEADER}' → column {openpyxl.utils.get_column_letter(qvd_col)}")

    qvd_map = {}
    for row_num in range(3, ws_lineage.max_row + 1):
        val = ws_lineage.cell(row=row_num, column=qvd_col).value
        if val and str(val).strip():
            qvd_map[row_num] = str(val).strip()

    print(f"         {len(qvd_map)} QVD name(s) found")

    # ── Step 2: Build QVD List lookup ─────────────────────────────────
    print(f"\nStep 2: Building lookup from '{QVDLIST_SHEET}'...")
    qvd_lookup = build_qvd_lookup(ws_qvdlist)

    # ── Step 3: Pre-read all QVW scripts ──────────────────────────────
    print(f"\nStep 3: Scanning QVW folder...")
    all_qvw = list(Path(QVW_FOLDER).rglob("*.qvw"))
    print(f"         {len(all_qvw)} QVW file(s) found. Reading scripts...")

    qvw_scripts = {}
    for qvw_path in all_qvw:
        script = extract_script_from_qvw(qvw_path)
        qvw_scripts[qvw_path] = script
        print(f"  ✓ {qvw_path.name} ({len(script):,} chars)")

    # ── Step 4: Match QVD → QVD List path → QVW → tables ─────────────
    print(f"\nStep 4: Matching QVD names to QVW scripts...")
    qvdlist_results = {}  # {qvdlist_row_num: {qvw_files, table_names}}

    for lin_row, qvd_name in qvd_map.items():
        qvd_stem = qvd_name.replace(".qvd", "").strip()
        lookup_entry = qvd_lookup.get(qvd_stem.lower())

        if not lookup_entry:
            print(f"  Row {lin_row}: {qvd_name} → [not in QVD List, skipping]")
            continue

        _, resolved_path, qvdlist_row = lookup_entry
        print(f"  Row {lin_row}: {qvd_name} → {resolved_path}")

        matching_qvws = []
        all_tables = set()

        for qvw_path, script_text in qvw_scripts.items():
            if script_creates_qvd(script_text, qvd_stem):
                matching_qvws.append(qvw_path.name)
                all_tables.update(parse_table_names(script_text))

        qvw_str    = " | ".join(matching_qvws) if matching_qvws else "NOT FOUND"
        tables_str = ", ".join(sorted(all_tables)) if all_tables else ""

        print(f"           QVW    : {qvw_str}")
        print(f"           Tables : {tables_str[:100]}{'...' if len(tables_str) > 100 else ''}")

        qvdlist_results[qvdlist_row] = {"qvw_files": qvw_str, "table_names": tables_str}

    # ── Step 5: Write results into QVD List only ──────────────────────
    print(f"\nStep 5: Writing results to '{QVDLIST_SHEET}'...")

    qvw_col    = ws_qvdlist.max_column + 1
    tables_col = ws_qvdlist.max_column + 2

    # Headers in row 1
    ws_qvdlist.cell(row=1, column=qvw_col).value    = "QVW File(s)"
    ws_qvdlist.cell(row=1, column=tables_col).value = "Source Table Names"

    for row_num, info in qvdlist_results.items():
        ws_qvdlist.cell(row=row_num, column=qvw_col).value    = info["qvw_files"]
        ws_qvdlist.cell(row=row_num, column=tables_col).value = info["table_names"]

    qvw_letter    = openpyxl.utils.get_column_letter(qvw_col)
    tables_letter = openpyxl.utils.get_column_letter(tables_col)
    print(f"         Column {qvw_letter} → QVW File(s)")
    print(f"         Column {tables_letter} → Source Table Names")

    # Save
    wb.save(OUTPUT_PATH)
    print(f"\nSaved → {OUTPUT_PATH}")
    print("Done.")


if __name__ == "__main__":
    main()
