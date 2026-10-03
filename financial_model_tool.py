"""
Strands tool: convert a PIM-type financial-model Excel file into a single
consolidated markdown file.

Per Praneet/Evan spec:
  - KEEP the original simplistic excel->markdown path untouched (not in this file).
  - This is the NEW "financial_model" path. It:
      1. Checks the file name against a list of known financial models (PIM/PDM).
         Only PIM is implemented for now; the list is trivially extensible.
      2. For a PIM, extracts the 5 standard tables from the `PortfolioView_Alt`
         sheet, located by header text (robust to variable row counts across
         submissions):
            - Cash Flow Analysis - [portfolio name]
            - Ratio Analysis
            - ICR Analysis - NAB
            - ICR Analysis - Combined Debt
            - LVR Analysis - Market Accepted Value: [value]
      3. Writes ONE markdown file containing all extracted tables and returns
         its path.

Scenario-sheet tables (Senior Debt / Combined Debt Live-v-Stored) are
intentionally skipped for now — add them to SCENARIO_ANCHORS when the layout
is confirmed and set EXTRACT_SCENARIO = True.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from strands import tool


# --------------------------------------------------------------------------- #
# Known model types. Extend this as new financial models need to be handled.
# Each entry: canonical name -> list of case-insensitive filename substrings.
# --------------------------------------------------------------------------- #
KNOWN_MODELS: dict[str, list[str]] = {
    "PIM": ["pim", "property investment model", "property_investment_model"],
    "PDM": ["pdm", "property development model", "property_development_model"],
}

IMPLEMENTED_MODELS = {"PIM"}  # focus on PIM for now


# --------------------------------------------------------------------------- #
# PIM / PortfolioView_Alt table anchors.
# Each table is found by scanning column A..B for a cell whose text STARTS WITH
# the anchor. Trailing "[enter portfolio name]" etc. is tolerated and the real
# filled-in suffix (if any) is kept as part of the title.
# --------------------------------------------------------------------------- #
SHEET_NAME = "PortfolioView_Alt"

PIM_ANCHORS: list[str] = [
    "Cash Flow Analysis",
    "Ratio Analysis",
    "ICR Analysis - NAB",
    "ICR Analysis - Combined Debt",
    "LVR Analysis",
]

# Flip on + fill in once the Scenario sheet layout is confirmed.
EXTRACT_SCENARIO = False
SCENARIO_ANCHORS: list[str] = [
    # "Senior Debt Ratios: Live Case",
    # "Combined Debt Ratios: Live Case",
]


@dataclass
class ExtractedTable:
    title: str
    df: pd.DataFrame
    start_row: int  # 1-based sheet row of the title


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def identify_model(file_path: str) -> Optional[str]:
    """Return the canonical model name if the filename matches a known model."""
    name = os.path.basename(file_path).lower()
    for model, needles in KNOWN_MODELS.items():
        if any(n in name for n in needles):
            return model
    return None


def _row_is_blank(ws: Worksheet, row: int, c_start: int, c_end: int) -> bool:
    for c in range(c_start, c_end + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and str(v).strip() != "":
            return False
    return True


def _find_anchor_rows(ws: Worksheet, anchors: list[str],
                      scan_cols: int = 3) -> list[tuple[str, str, int]]:
    """
    Locate each anchor by header text.

    Returns list of (anchor, full_title_as_in_cell, row) sorted by row.
    Normalises hyphen/dash and whitespace so 'ICR Analysis - NAB' matches
    'ICR Analysis – NAB' etc.
    """
    def norm(s: str) -> str:
        s = str(s).replace("\u2013", "-").replace("\u2014", "-")
        return re.sub(r"\s+", " ", s).strip().lower()

    found: list[tuple[str, str, int]] = []
    anchor_norm = {a: norm(a) for a in anchors}
    remaining = set(anchors)

    for row in range(1, ws.max_row + 1):
        if not remaining:
            break
        for col in range(1, scan_cols + 1):
            cell = ws.cell(row=row, column=col).value
            if cell is None:
                continue
            cn = norm(cell)
            for a in list(remaining):
                if cn.startswith(anchor_norm[a]):
                    found.append((a, str(cell).strip(), row))
                    remaining.discard(a)
                    break

    missing = [a for a in anchors if a not in {f[0] for f in found}]
    if missing:
        raise ValueError(
            f"Could not locate table header(s) in '{ws.title}': {missing}. "
            f"Headers found: {[f[1] for f in found]}"
        )

    found.sort(key=lambda t: t[2])
    return found


def _extract_block(ws: Worksheet, title_row: int, next_title_row: Optional[int],
                   c_start: int = 1, c_end: Optional[int] = None) -> pd.DataFrame:
    """
    Read the block between a title row and the next title (or next blank gap),
    into a DataFrame. First non-blank row after the title is treated as headers.
    """
    if c_end is None:
        c_end = ws.max_column

    body_start = title_row + 1
    # upper bound: just before the next title, else end of sheet
    hard_end = (next_title_row - 1) if next_title_row else ws.max_row

    # walk down, stop at the FIRST blank row that's followed by another blank
    # (a single blank inside a table is tolerated; a double blank ends it)
    rows: list[int] = []
    r = body_start
    blanks = 0
    while r <= hard_end:
        if _row_is_blank(ws, r, c_start, c_end):
            blanks += 1
            if blanks >= 2:
                break
        else:
            blanks = 0
            rows.append(r)
        r += 1

    if not rows:
        return pd.DataFrame()

    # trim trailing all-empty columns by finding max used col in this block
    used_end = c_start
    for rr in rows:
        for cc in range(c_end, c_start - 1, -1):
            v = ws.cell(row=rr, column=cc).value
            if v is not None and str(v).strip() != "":
                used_end = max(used_end, cc)
                break

    data = []
    for rr in rows:
        data.append([ws.cell(row=rr, column=cc).value
                     for cc in range(c_start, used_end + 1)])

    df = pd.DataFrame(data)
    # first row as header
    df.columns = [("" if v is None else str(v).strip()) for v in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    # drop fully empty columns
    df = df.loc[:, ~(df.isna() | (df.astype(str).apply(lambda s: s.str.strip()) == "")).all()]
    # blank out remaining NaNs for clean markdown
    df = df.fillna("")
    return df


def _extract_tables(ws: Worksheet, anchors: list[str]) -> list[ExtractedTable]:
    located = _find_anchor_rows(ws, anchors)
    out: list[ExtractedTable] = []
    for i, (_, title, row) in enumerate(located):
        next_row = located[i + 1][2] if i + 1 < len(located) else None
        df = _extract_block(ws, row, next_row)
        out.append(ExtractedTable(title=title, df=df, start_row=row))
    return out


def _tables_to_markdown(model: str, source: str,
                        tables: list[ExtractedTable]) -> str:
    parts = [f"# {model} Extract — {os.path.basename(source)}", ""]
    for t in tables:
        parts.append(f"## {t.title}")
        parts.append("")
        if t.df.empty:
            parts.append("_(no data found)_")
        else:
            parts.append(t.df.to_markdown(index=False))
        parts.append("")
    return "\n".join(parts)


# --------------------------------------------------------------------------- #
# The Strands tool
# --------------------------------------------------------------------------- #
@tool
def financial_model_to_markdown(file_path: str,
                                output_dir: Optional[str] = None) -> str:
    """Convert a known financial-model Excel file (currently PIM) into a single
    consolidated markdown file and return the path to that file.

    This is the specialised path for structured financial models. It checks the
    filename against known model types (PIM/PDM), and for a PIM extracts the 5
    standard tables from the PortfolioView_Alt sheet (Cash Flow Analysis, Ratio
    Analysis, ICR Analysis - NAB, ICR Analysis - Combined Debt, LVR Analysis),
    writing them all into one markdown file. Tables are located by header text,
    so variable row counts across submissions are handled automatically.

    Args:
        file_path: Path to the .xlsx/.xlsm financial-model file.
        output_dir: Directory to write the markdown into. Defaults to the
            source file's directory.

    Returns:
        The absolute path to the written markdown file.

    Raises:
        FileNotFoundError: if file_path does not exist.
        ValueError: if the file is not a recognised/implemented model, the
            target sheet is missing, or an expected table header is not found.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(file_path)

    model = identify_model(file_path)
    if model is None:
        raise ValueError(
            f"'{os.path.basename(file_path)}' does not match any known financial "
            f"model {list(KNOWN_MODELS)}. Use the simplistic excel->markdown path instead."
        )
    if model not in IMPLEMENTED_MODELS:
        raise ValueError(f"Model '{model}' recognised but not implemented yet.")

    wb = load_workbook(file_path, data_only=True, read_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{SHEET_NAME}' not found. Sheets: {wb.sheetnames}"
        )

    tables = _extract_tables(wb[SHEET_NAME], PIM_ANCHORS)

    if EXTRACT_SCENARIO and SCENARIO_ANCHORS:
        if "Scenario" in wb.sheetnames:
            tables += _extract_tables(wb["Scenario"], SCENARIO_ANCHORS)

    md = _tables_to_markdown(model, file_path, tables)

    out_dir = output_dir or os.path.dirname(os.path.abspath(file_path))
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(file_path))[0]
    out_path = os.path.join(out_dir, f"{stem}.md")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md)

    return os.path.abspath(out_path)
