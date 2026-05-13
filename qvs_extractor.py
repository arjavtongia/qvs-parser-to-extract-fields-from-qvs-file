import re
import csv
import sys
from pathlib import Path

# =========================================================
# ADVANCED QVS PARSER
# =========================================================

TAB_REGEX = re.compile(r'^\s*///\$tab\s+(.*)', re.IGNORECASE)

TABLE_REGEX = re.compile(
    r'^\s*([A-Za-z0-9_]+)\s*:\s*$',
    re.IGNORECASE
)

JOIN_REGEX = re.compile(
    r'(LEFT|RIGHT|INNER|OUTER|FULL)?\s*JOIN\s*\(\s*([^)]+)\s*\)',
    re.IGNORECASE
)

LOAD_START_REGEX = re.compile(
    r'\bLOAD\b',
    re.IGNORECASE
)

SQL_SELECT_REGEX = re.compile(
    r'\bSQL\s+SELECT\b',
    re.IGNORECASE
)

RESIDENT_REGEX = re.compile(
    r'RESIDENT\s+([A-Za-z0-9_]+)',
    re.IGNORECASE
)

INLINE_REGEX = re.compile(
    r'INLINE\s*\[',
    re.IGNORECASE
)

# =========================================================
# REMOVE COMMENTS
# =========================================================

def remove_comments(text):

    # Remove block comments
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)

    cleaned_lines = []

    for line in text.splitlines():

        stripped = line.strip()

        # Skip single line comments
        if stripped.startswith("//"):
            continue

        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)

# =========================================================
# EXTRACT TAB NAME
# =========================================================

def extract_tab_name(line, current_tab):

    match = TAB_REGEX.match(line)

    if match:
        return match.group(1).strip()

    return current_tab

# =========================================================
# EXTRACT TABLE NAME
# =========================================================

def extract_table_name(line, current_table):

    table_match = TABLE_REGEX.match(line)

    if table_match:
        return table_match.group(1).strip()

    join_match = JOIN_REGEX.search(line)

    if join_match:
        return join_match.group(2).strip()

    return current_table

# =========================================================
# SAFE FIELD SPLITTER
# Handles commas inside functions/brackets/quotes
# =========================================================

def split_fields_safely(text):

    fields = []

    current = []

    round_depth = 0
    square_depth = 0

    in_single_quote = False
    in_double_quote = False

    i = 0

    while i < len(text):

        ch = text[i]

        # -----------------------------
        # QUOTES
        # -----------------------------

        if ch == "'" and not in_double_quote:
            in_single_quote = not in_single_quote
            current.append(ch)
            i += 1
            continue

        if ch == '"' and not in_single_quote:
            in_double_quote = not in_double_quote
            current.append(ch)
            i += 1
            continue

        # -----------------------------
        # BRACKET DEPTH
        # -----------------------------

        if not in_single_quote and not in_double_quote:

            if ch == '(':
                round_depth += 1

            elif ch == ')':
                round_depth -= 1

            elif ch == '[':
                square_depth += 1

            elif ch == ']':
                square_depth -= 1

            # -----------------------------
            # TOP LEVEL COMMA
            # -----------------------------

            elif (
                ch == ','
                and round_depth == 0
                and square_depth == 0
            ):

                field = ''.join(current).strip()

                if field:
                    fields.append(field)

                current = []

                i += 1
                continue

        current.append(ch)

        i += 1

    # Final field
    final_field = ''.join(current).strip()

    if final_field:
        fields.append(final_field)

    return fields

# =========================================================
# EXTRACT LOAD FIELDS
# =========================================================

def extract_fields(load_block):

    fields = []

    lines = load_block.splitlines()

    collecting = False

    collected_text = []

    for line in lines:

        stripped = line.strip()

        if not stripped:
            continue

        # -----------------------------
        # START COLLECTING AFTER LOAD
        # -----------------------------

        if re.search(r'\bLOAD\b', stripped, re.IGNORECASE):

            collecting = True

            stripped = re.sub(
                r'.*?\bLOAD\b',
                '',
                stripped,
                flags=re.IGNORECASE
            ).strip()

            if stripped:
                collected_text.append(stripped)

            continue

        # -----------------------------
        # COLLECT FIELD LINES
        # -----------------------------

        if collecting:

            upper = stripped.upper()

            # Stop collection
            if (
                upper.startswith("FROM")
                or upper.startswith("RESIDENT")
                or upper.startswith("INLINE")
                or upper.startswith("WHERE")
                or upper.startswith("GROUP BY")
                or upper.startswith("ORDER BY")
                or upper.startswith("SQL")
            ):
                break

            collected_text.append(stripped)

    # Combine all field text
    full_text = ' '.join(collected_text)

    # Split safely
    split_result = split_fields_safely(full_text)

    for fld in split_result:

        fld = fld.strip()

        fld = re.sub(r',$', '', fld)

        fld = fld.strip()

        if fld:
            fields.append(fld)

    return fields

# =========================================================
# EXTRACT SQL FIELDS
# =========================================================

def extract_sql_fields(block):

    fields = []

    collecting = False

    collected_text = []

    for line in block.splitlines():

        stripped = line.strip()

        if re.search(r'SQL\s+SELECT', stripped, re.IGNORECASE):

            collecting = True

            stripped = re.sub(
                r'SQL\s+SELECT',
                '',
                stripped,
                flags=re.IGNORECASE
            ).strip()

            if stripped:
                collected_text.append(stripped)

            continue

        if collecting:

            if stripped.upper().startswith("FROM"):
                break

            collected_text.append(stripped)

    full_text = ' '.join(collected_text)

    split_result = split_fields_safely(full_text)

    for fld in split_result:

        fld = fld.strip()

        fld = re.sub(r',$', '', fld)

        fld = fld.strip()

        if fld:
            fields.append(fld)

    return fields

# =========================================================
# EXTRACT FROM
# =========================================================

def extract_from(block):

    lines = [x.strip() for x in block.splitlines() if x.strip()]

    # -----------------------------------------------------
    # RESIDENT
    # -----------------------------------------------------

    for line in lines:

        resident_match = RESIDENT_REGEX.search(line)

        if resident_match:
            return f"RESIDENT {resident_match.group(1)}"

    # -----------------------------------------------------
    # INLINE
    # -----------------------------------------------------

    for line in lines:

        if INLINE_REGEX.search(line):
            return "INLINE"

    # -----------------------------------------------------
    # SQL FROM
    # -----------------------------------------------------

    sql_mode = False

    for line in lines:

        if re.search(r'SQL\s+SELECT', line, re.IGNORECASE):
            sql_mode = True

        if sql_mode and line.upper().startswith("FROM"):

            from_val = re.sub(
                r'^FROM',
                '',
                line,
                flags=re.IGNORECASE
            ).strip()

            from_val = from_val.rstrip(';')

            return from_val

    # -----------------------------------------------------
    # NORMAL FROM
    # -----------------------------------------------------

    for idx, line in enumerate(lines):

        if line.upper().startswith("FROM"):

            from_lines = []

            first = re.sub(
                r'^FROM',
                '',
                line,
                flags=re.IGNORECASE
            ).strip()

            if first:
                from_lines.append(first)

            j = idx + 1

            while j < len(lines):

                nxt = lines[j]

                upper = nxt.upper()

                if (
                    upper.startswith("WHERE")
                    or upper.startswith("GROUP BY")
                    or upper.startswith("ORDER BY")
                    or upper.startswith("LOAD")
                    or upper.startswith("SQL")
                ):
                    break

                from_lines.append(nxt)

                # Stop after qualifier block
                if ")" in nxt:
                    break

                j += 1

            result = " ".join(from_lines)

            result = result.replace("[", "").replace("]", "")

            result = result.rstrip(";")

            return result.strip()

    return ""

# =========================================================
# SPLIT INTO BLOCKS
# =========================================================

def split_blocks(content):

    blocks = []

    current = []

    for line in content.splitlines():

        current.append(line)

        if ';' in line:

            blocks.append("\n".join(current))

            current = []

    if current:
        blocks.append("\n".join(current))

    return blocks

# =========================================================
# MAIN PARSER
# =========================================================

def parse_qvs(filepath):

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    content = remove_comments(content)

    blocks = split_blocks(content)

    rows = []

    seen = set()

    current_tab = "MAIN"
    current_table = "UNNAMED_TABLE"

    for block in blocks:

        lines = block.splitlines()

        # -------------------------------------------------
        # Update tab/table context
        # -------------------------------------------------

        for line in lines:

            current_tab = extract_tab_name(
                line,
                current_tab
            )

            current_table = extract_table_name(
                line,
                current_table
            )

        fields = []

        # -------------------------------------------------
        # LOAD fields
        # -------------------------------------------------

        if LOAD_START_REGEX.search(block):
            fields.extend(extract_fields(block))

        # -------------------------------------------------
        # SQL fields
        # -------------------------------------------------

        if SQL_SELECT_REGEX.search(block):
            fields.extend(extract_sql_fields(block))

        if not fields:
            continue

        from_value = extract_from(block)

        # -------------------------------------------------
        # Add rows
        # -------------------------------------------------

        for field in fields:

            row = (
                current_tab,
                current_table,
                field,
                from_value
            )

            if row not in seen:

                seen.add(row)

                rows.append([
                    current_tab,
                    current_table,
                    field,
                    from_value
                ])

    return rows

# =========================================================
# WRITE CSV
# =========================================================

def write_csv(rows, output_file):

    with open(output_file, 'w', newline='', encoding='utf-8') as f:

        writer = csv.writer(f)

        writer.writerow([
            "Tab Name",
            "Table Name",
            "Field",
            "From"
        ])

        writer.writerows(rows)

# =========================================================
# MAIN
# =========================================================

def main():

    if len(sys.argv) != 2:

        print("\nUsage:")
        print("python script.py input.qvs")

        sys.exit(1)

    input_file = Path(sys.argv[1])

    if not input_file.exists():

        print("\nFile not found.")
        sys.exit(1)

    output_file = input_file.with_suffix(".csv")

    rows = parse_qvs(input_file)

    write_csv(rows, output_file)

    print(f"\nCSV created successfully:")
    print(output_file)

    print(f"\nTotal rows extracted: {len(rows)}")

if __name__ == "__main__":
    main()
