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

    # remove block comments
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)

    cleaned_lines = []

    for line in text.splitlines():

        stripped = line.strip()

        # remove // comments
        if stripped.startswith("//"):
            continue

        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


# =========================================================
# EXTRACT TABS
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
# CLEAN FIELD
# =========================================================

def clean_field(field):

    field = field.strip()

    field = re.sub(r',$', '', field)

    return field.strip()


# =========================================================
# EXTRACT FIELDS
# =========================================================

def extract_fields(load_block):

    fields = []

    lines = load_block.splitlines()

    collecting = False

    for line in lines:

        stripped = line.strip()

        if not stripped:
            continue

        # start collecting after LOAD
        if re.search(r'\bLOAD\b', stripped, re.IGNORECASE):

            collecting = True

            stripped = re.sub(
                r'.*?\bLOAD\b',
                '',
                stripped,
                flags=re.IGNORECASE
            ).strip()

            if stripped:
                line_fields = stripped.split(',')

                for fld in line_fields:
                    cleaned = clean_field(fld)

                    if cleaned:
                        fields.append(cleaned)

            continue

        if collecting:

            upper = stripped.upper()

            # stop conditions
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

            line_fields = stripped.split(',')

            for fld in line_fields:

                cleaned = clean_field(fld)

                if cleaned:
                    fields.append(cleaned)

    return fields


# =========================================================
# EXTRACT SQL FIELDS
# =========================================================

def extract_sql_fields(block):

    fields = []

    collecting = False

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
                for fld in stripped.split(','):
                    fld = clean_field(fld)

                    if fld:
                        fields.append(fld)

            continue

        if collecting:

            if stripped.upper().startswith("FROM"):
                break

            for fld in stripped.split(','):

                fld = clean_field(fld)

                if fld:
                    fields.append(fld)

    return fields


# =========================================================
# EXTRACT FROM
# =========================================================

def extract_from(block):

    lines = [x.strip() for x in block.splitlines() if x.strip()]

    # RESIDENT
    for line in lines:

        resident_match = RESIDENT_REGEX.search(line)

        if resident_match:
            return f"RESIDENT {resident_match.group(1)}"

    # INLINE
    for line in lines:

        if INLINE_REGEX.search(line):
            return "INLINE"

    # SQL FROM
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

    # NORMAL FROM
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

            # capture multiline FROM
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

                # stop after qualifier
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

        if LOAD_START_REGEX.search(block):
            fields.extend(extract_fields(block))

        if SQL_SELECT_REGEX.search(block):
            fields.extend(extract_sql_fields(block))

        if not fields:
            continue

        from_value = extract_from(block)

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

        print("Usage:")
        print("python script.py input.qvs")

        sys.exit(1)

    input_file = Path(sys.argv[1])

    if not input_file.exists():

        print("File not found.")
        sys.exit(1)

    output_file = input_file.with_suffix(".csv")

    rows = parse_qvs(input_file)

    write_csv(rows, output_file)

    print(f"\nCSV created:")
    print(output_file)

    print(f"\nTotal rows: {len(rows)}")


if __name__ == "__main__":
    main()
