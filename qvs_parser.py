"""
QVS Script Parser
Extracts: Tab Name, Table Name, Fields, FROM source
Usage: python qvs_parser.py <script.qvs>
       python qvs_parser.py <script.qvs> --output results.csv
"""

import re
import sys
import csv
import json
from dataclasses import dataclass, field, asdict
from typing import Optional


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class LoadBlock:
    tab_name: str
    table_name: Optional[str]
    load_type: str          # LOAD | MAPPING LOAD | PRECEDING LOAD
    fields: list[str]
    from_source: Optional[str]
    autogenerate: Optional[str]
    raw_snippet: str = ""   # first ~120 chars for debugging


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Strip inline comments  e.g.  MyField,  //GB: PK
_INLINE_COMMENT = re.compile(r"\s*//.*$", re.MULTILINE)

def strip_inline_comments(text: str) -> str:
    return _INLINE_COMMENT.sub("", text)


def clean_field(raw: str) -> str:
    """Normalise a single field expression."""
    f = raw.strip().rstrip(",").strip()
    f = re.sub(r"\s+", " ", f)
    return f


def parse_field_list(field_block: str) -> list[str]:
    """
    Split a raw field block (text between LOAD and FROM/AutoGenerate/;)
    into individual field expressions, respecting nested parentheses.
    """
    fields = []
    depth = 0
    current = []

    for ch in field_block:
        if ch == "(":
            depth += 1
            current.append(ch)
        elif ch == ")":
            depth -= 1
            current.append(ch)
        elif ch == "," and depth == 0:
            token = "".join(current).strip()
            if token:
                fields.append(clean_field(token))
            current = []
        else:
            current.append(ch)

    # Last token
    token = "".join(current).strip()
    if token:
        fields.append(clean_field(token))

    return [f for f in fields if f]


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

# Patterns
_TAB_RE       = re.compile(r"^/{2,3}\$tab\s+(.+)", re.IGNORECASE | re.MULTILINE)
_TABLE_LABEL  = re.compile(r"^([A-Za-z_][A-Za-z0-9_\s]*):\s*$", re.MULTILINE)
_LOAD_BLOCK   = re.compile(
    r"(MAPPING\s+LOAD|LOAD)\s+"     # LOAD keyword (with optional MAPPING prefix)
    r"(.*?)"                         # field list
    r"(?=\bFROM\b|\bAutoGenerate\b|\bWHERE\b|;)",
    re.IGNORECASE | re.DOTALL
)
_FROM_RE      = re.compile(
    r"\bFROM\b\s+([^\s;]+(?:\s+\([^)]+\))?)",
    re.IGNORECASE
)
_AUTOGEN_RE   = re.compile(r"\bAutoGenerate\b\s+(\S+)", re.IGNORECASE)
_PRECEDING_RE = re.compile(r"^\s*LOAD\b", re.IGNORECASE | re.MULTILINE)


def parse_qvs(script: str) -> list[LoadBlock]:
    """Parse a full QVS script and return a list of LoadBlock objects."""

    # ------------------------------------------------------------------
    # 1. Split into tab sections using ///$tab markers
    # ------------------------------------------------------------------
    tab_positions = [(m.start(), m.group(1).strip()) for m in _TAB_RE.finditer(script)]

    # Build list of (start, end, tab_name) segments
    segments: list[tuple[int, int, str]] = []
    for i, (pos, tab) in enumerate(tab_positions):
        end = tab_positions[i + 1][0] if i + 1 < len(tab_positions) else len(script)
        segments.append((pos, end, tab))

    # If no ///$tab markers at all, treat the whole file as one anonymous tab
    if not segments:
        segments = [(0, len(script), "UNKNOWN")]

    results: list[LoadBlock] = []

    for seg_start, seg_end, tab_name in segments:
        seg_text = script[seg_start:seg_end]

        # Find all LOAD blocks within this tab segment
        for load_match in _LOAD_BLOCK.finditer(seg_text):
            load_keyword = load_match.group(1).strip().upper()
            raw_fields   = load_match.group(2)

            # ---- load type ------------------------------------------------
            if "MAPPING" in load_keyword:
                load_type = "MAPPING LOAD"
            else:
                load_type = "LOAD"

            # ---- Check for preceding LOAD ---------------------------------
            # A preceding LOAD is a LOAD that has no FROM; the previous LOAD
            # acts as its source. We detect this post-hoc (see below).

            # ---- Fields ---------------------------------------------------
            cleaned_fields_text = strip_inline_comments(raw_fields)
            fields = parse_field_list(cleaned_fields_text)

            # ---- FROM / AutoGenerate (look ahead from end of LOAD block) --
            # Limit look-ahead to the current statement (up to next bare ';')
            look_ahead_start = load_match.end()
            rest_of_seg      = seg_text[look_ahead_start:]
            # Find the closing semicolon of this LOAD statement
            semi_match = re.search(r"(?<!\w);", rest_of_seg)
            if semi_match:
                look_ahead_end = look_ahead_start + semi_match.end()
            else:
                look_ahead_end = min(look_ahead_start + 800, seg_end)
            look_ahead_text  = seg_text[look_ahead_start:look_ahead_end]

            from_source  = None
            autogenerate = None

            from_m = _FROM_RE.search(look_ahead_text)
            if from_m:
                from_source = from_m.group(1).strip()

            ag_m = _AUTOGEN_RE.search(look_ahead_text)
            if ag_m:
                autogenerate = ag_m.group(1).strip()

            # ---- Table label (look BEFORE the LOAD keyword) ---------------
            text_before_load = seg_text[:load_match.start()]
            table_name = None
            for lbl_m in _TABLE_LABEL.finditer(text_before_load):
                candidate = lbl_m.group(1).strip()
                # Skip QVS keywords that might appear as labels
                if candidate.upper() not in {"LOAD", "FROM", "WHERE", "MAPPING",
                                              "SET", "LET", "TRACE", "CALL"}:
                    table_name = candidate
            # table_name is now the last valid label before this LOAD block

            raw_snippet = seg_text[max(0, load_match.start()-30):load_match.start()+90]
            raw_snippet = raw_snippet.replace("\n", " ").strip()

            block = LoadBlock(
                tab_name=tab_name,
                table_name=table_name,
                load_type=load_type,
                fields=fields,
                from_source=from_source,
                autogenerate=autogenerate,
                raw_snippet=raw_snippet[:120],
            )
            results.append(block)

    # ------------------------------------------------------------------
    # 2. Post-process: flag preceding LOADs (LOAD with no FROM that is
    #    immediately followed by another LOAD for the same table)
    # ------------------------------------------------------------------
    for i, blk in enumerate(results):
        if blk.from_source is None and blk.autogenerate is None:
            # No source — this is a preceding LOAD
            blk.load_type = "PRECEDING LOAD"

    return results


# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

def print_table(blocks: list[LoadBlock]) -> None:
    sep = "-" * 100
    for b in blocks:
        print(sep)
        print(f"  Tab       : {b.tab_name}")
        print(f"  Table     : {b.table_name or '(anonymous)'}")
        print(f"  Load Type : {b.load_type}")
        if b.from_source:
            print(f"  FROM      : {b.from_source}")
        if b.autogenerate:
            print(f"  AutoGen   : {b.autogenerate} rows")
        print(f"  Fields ({len(b.fields)}):")
        for f in b.fields:
            print(f"    • {f}")
    print(sep)
    print(f"\nTotal LOAD blocks found: {len(blocks)}")


def export_csv(blocks: list[LoadBlock], path: str) -> None:
    """One row per field for easy analysis in Excel / Databricks."""
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["tab_name", "table_name", "load_type",
                         "field_expression", "from_source", "autogenerate"])
        for b in blocks:
            if b.fields:
                for f in b.fields:
                    writer.writerow([b.tab_name, b.table_name or "",
                                     b.load_type, f,
                                     b.from_source or "", b.autogenerate or ""])
            else:
                writer.writerow([b.tab_name, b.table_name or "",
                                 b.load_type, "",
                                 b.from_source or "", b.autogenerate or ""])
    print(f"CSV written → {path}")


def export_json(blocks: list[LoadBlock], path: str) -> None:
    data = []
    for b in blocks:
        d = asdict(b)
        d.pop("raw_snippet", None)
        data.append(d)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    print(f"JSON written → {path}")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Parse a QlikView/Qlik Sense .qvs script and extract "
                    "tab names, table names, fields, and FROM sources."
    )
    parser.add_argument("script", help="Path to the .qvs file")
    parser.add_argument("--output", "-o",
                        help="Optional output file path. "
                             "Extension determines format: .csv or .json")
    parser.add_argument("--encoding", default="utf-8",
                        help="File encoding (default: utf-8)")
    args = parser.parse_args()

    with open(args.script, "r", encoding=args.encoding, errors="replace") as fh:
        script_text = fh.read()

    blocks = parse_qvs(script_text)

    print_table(blocks)

    if args.output:
        ext = args.output.rsplit(".", 1)[-1].lower()
        if ext == "csv":
            export_csv(blocks, args.output)
        elif ext == "json":
            export_json(blocks, args.output)
        else:
            print(f"Unknown extension '.{ext}' — defaulting to CSV.")
            export_csv(blocks, args.output)


if __name__ == "__main__":
    main()
