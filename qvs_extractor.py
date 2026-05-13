import re
import csv
import sys
from pathlib import Path

# =========================================================
# QVS Parser
# Extracts:
# Tab Name | Table Name | Field | From
# =========================================================

TAB_PATTERN = re.compile(r'^\s*///\$tab\s+(.*)', re.IGNORECASE)
TABLE_LABEL_PATTERN = re.compile(r'^\s*([A-Za-z0-9_\-\s]+)\s*:\s*$')
JOIN_PATTERN = re.compile(
    r'^\s*(LEFT|RIGHT|INNER|OUTER|FULL)?\s*JOIN\s*\(\s*([^)]+)\s*\)',
    re.IGNORECASE
)

FROM_PATTERN = re.compile(
    r'FROM\s+\[(.*?)\]\s*(\([^)]+\))?',
    re.IGNORECASE | re.DOTALL
)

FROM_NO_BRACKET_PATTERN = re.compile(
    r'FROM\s+([^\s;]+)\s*(\([^)]+\))?',
    re.IGNORECASE | re.DOTALL
)

RESIDENT_PATTERN = re.compile(
    r'RESIDENT\s+([A-Za-z0-9_]+)',
    re.IGNORECASE
)

SQL_FROM_PATTERN = re.compile(
    r'FROM\s+([A-Za-z0-9_.]+)',
    re.IGNORECASE
)

INLINE_PATTERN = re.compile(
    r'INLINE\s*\[',
    re.IGNORECASE
)

LOAD_PATTERN = re.compile(
    r'\bLOAD\b',
    re.IGNORECASE
)

SQL_SELECT_PATTERN = re.compile(
    r'\bSQL\s+SELECT\b',
    re.IGNORECASE
)


def remove_comments(text):
    """
    Remove:
    - // comments
    - /* */ comments
    """
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)
    text = re.sub(r'//.*', '', text)
    return text


def split_statements(text):
    """
    Split QVS script into statements using semicolon.
    """
    statements = []
    current = []

    for line in text.splitlines():
        current.append(line)

        if ';' in line:
            statements.append('\n'.join(current))
            current = []

    if current:
        statements.append('\n'.join(current))

    return statements


def clean_field(field):
    field = field.strip()

    # Remove trailing commas
    field = re.sub(r',$', '', field)

    return field.strip()


def extract_load_fields(statement):
    """
    Extract fields from LOAD blocks.
    """
    fields = []

    load_matches = list(re.finditer(r'\bLOAD\b', statement, re.IGNORECASE))

    for idx, match in enumerate(load_matches):

        start = match.end()

        remaining = statement[start:]

        stop_patterns = [
            r'\bFROM\b',
            r'\bRESIDENT\b',
            r'\bINLINE\b',
            r'\bSQL\b',
            r';'
        ]

        stop_positions = []

        for pat in stop_patterns:
            m = re.search(pat, remaining, re.IGNORECASE)
            if m:
                stop_positions.append(m.start())

        if stop_positions:
            end_pos = min(stop_positions)
            load_block = remaining[:end_pos]
        else:
            load_block = remaining

        raw_fields = load_block.split(',')

        for field in raw_fields:
            cleaned = clean_field(field)

            if cleaned:
                fields.append(cleaned)

    return fields


def extract_sql_fields(statement):
    """
    Extract fields from SQL SELECT blocks.
    """
    fields = []

    sql_match = re.search(
        r'SQL\s+SELECT(.*?)FROM',
        statement,
        re.IGNORECASE | re.DOTALL
    )

    if sql_match:
        field_block = sql_match.group(1)

        raw_fields = field_block.split(',')

        for field in raw_fields:
            cleaned = clean_field(field)

            if cleaned:
                fields.append(cleaned)

    return fields


def extract_from(statement):
    """
    Extract FROM source.
    """

    # RESIDENT
    resident_match = RESIDENT_PATTERN.search(statement)
    if resident_match:
        return f"RESIDENT {resident_match.group(1).strip()}"

    # INLINE
    if INLINE_PATTERN.search(statement):
        return "INLINE"

    # SQL FROM
    if SQL_SELECT_PATTERN.search(statement):
        sql_from = SQL_FROM_PATTERN.search(statement)

        if sql_from:
            return sql_from.group(1).strip()

    # Standard FROM [file]
    from_match = FROM_PATTERN.search(statement)

    if from_match:
        source = from_match.group(1).strip()
        qualifier = from_match.group(2) or ''

        return f"{source} {qualifier}".strip()

    # FROM without brackets
    from_match = FROM_NO_BRACKET_PATTERN.search(statement)

    if from_match:
        source = from_match.group(1).strip()
        qualifier = from_match.group(2) or ''

        return f"{source} {qualifier}".strip()

    return ""


def extract_table_name(statement, current_table_name):
    """
    Determine table name.
    """

    lines = statement.splitlines()

    for line in lines:

        # Explicit table label
        table_match = TABLE_LABEL_PATTERN.match(line)

        if table_match:
            return table_match.group(1).strip()

        # JOIN target
        join_match = JOIN_PATTERN.match(line)

        if join_match:
            return join_match.group(2).strip()

    return current_table_name or "UNNAMED_TABLE"


def extract_tab_name(statement, current_tab):
    """
    Extract tab name from ///$tab
    """
    lines = statement.splitlines()

    for line in lines:
        tab_match = TAB_PATTERN.match(line)

        if tab_match:
            return tab_match.group(1).strip()

    return current_tab


def process_qvs(qvs_path):
    with open(qvs_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    content = remove_comments(content)

    statements = split_statements(content)

    rows = []
    seen = set()

    current_tab = ""
    current_table = ""

    for statement in statements:

        current_tab = extract_tab_name(statement, current_tab)

        current_table = extract_table_name(statement, current_table)

        from_value = extract_from(statement)

        fields = []

        if LOAD_PATTERN.search(statement):
            fields.extend(extract_load_fields(statement))

        if SQL_SELECT_PATTERN.search(statement):
            fields.extend(extract_sql_fields(statement))

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


def write_csv(rows, output_path):

    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:

        writer = csv.writer(csvfile)

        writer.writerow([
            'Tab Name',
            'Table Name',
            'Field',
            'From'
        ])

        writer.writerows(rows)


def main():

    if len(sys.argv) != 2:
        print("Usage:")
        print("python script.py input.qvs")
        sys.exit(1)

    input_path = Path(sys.argv[1])

    if not input_path.exists():
        print(f"File not found: {input_path}")
        sys.exit(1)

    output_path = input_path.with_suffix('.csv')

    rows = process_qvs(input_path)

    write_csv(rows, output_path)

    print(f"\nCSV generated successfully:")
    print(output_path)
    print(f"\nTotal rows: {len(rows)}")


if __name__ == "__main__":
    main()
