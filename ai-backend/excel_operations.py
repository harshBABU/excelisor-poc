# aibackend/excel_operations.py

import re
from typing import Any, Dict, List, Optional, Tuple


# =====================
# Validation heuristics
# =====================

ADDRESS_RE = re.compile(r"^[A-Za-z]{1,3}[0-9]{1,7}$")  # e.g., A1, AA10, XFD1048576 (Excel max not enforced here)


def is_valid_cell_address(addr: str) -> bool:
    """Basic check for A1-style addresses (no ranges like A1:B2 here)."""
    if not isinstance(addr, str) or not addr:
        return False
    return bool(ADDRESS_RE.match(addr))


def normalize_formula_text(formula: Any) -> str:
    """
    Ensure Excel formulas begin with '=' and are strings.
    Keeps non-string values safe by stringifying.
    """
    if formula is None:
        return ""
    s = str(formula).strip()
    if not s:
        return ""
    return s if s.startswith("=") else f"={s}"


# ==========================
# Payload construction utils
# ==========================

def build_worksheet_payload(
    name: str,
    cells: Optional[List[Dict[str, Any]]] = None,
    formulas: Optional[List[Dict[str, Any]]] = None,
    notes: str = "",
    conditional_formatting: Optional[List[Dict[str, Any]]] = None,
    column_widths: Optional[List[Tuple[str, float]]] = None,
    freeze_panes: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Construct a normalized worksheet payload for the frontend to apply.

    Args:
      name: Worksheet name requested (frontend may create or reuse).
      cells: List of {"address":"A1","value":<scalar or string>}.
      formulas: List of {"address":"F2","formula":"=SUM(A2:A100)","description":"optional"}.
      notes: Optional message; when provided, a note header is added at A1/A2 before other cells.
      conditional_formatting: Optional styling hints:
        [
          {"range":"A1:C10","type":"colorScale","style":"green_to_red"},
          {"range":"D:D","type":"greaterThan","threshold":10,"style":"red_fill"}
        ]
        The frontend should apply these best-effort with try/catch.
      column_widths: Optional list of (column_letter, width) e.g., [("A", 20.0), ("B", 12.0)].
      freeze_panes: Optional address like "A2" to freeze rows above/columns left (frontend controlled).

    Returns:
      Dict with fields: name, cells, formulas, notes, conditional_formatting, column_widths, freeze_panes.
    """
    normalized_cells: List[Dict[str, Any]] = []
    normalized_formulas: List[Dict[str, Any]] = []

    # Notes block (optional) — put before other cells so it renders at the top
    if notes:
        normalized_cells.append({"address": "A1", "value": "AI Notes"})
        normalized_cells.append({"address": "A2", "value": notes})

    # Cells normalization
    for c in cells or []:
        addr = c.get("address") or c.get("cell") or ""
        val = c.get("value", "")
        # Best-effort validation; if invalid, skip rather than crash
        if isinstance(addr, str) and addr:
            normalized_cells.append({"address": addr, "value": val})

    # Formulas normalization
    for f in formulas or []:
        addr = f.get("address") or f.get("cell") or ""
        form = normalize_formula_text(f.get("formula", ""))
        # Skip empty formulas; address can be arbitrary but should be present
        if isinstance(addr, str) and addr and form:
            entry = {"address": addr, "formula": form}
            desc = f.get("description")
            if isinstance(desc, str) and desc:
                entry["description"] = desc
            normalized_formulas.append(entry)

    payload = {
        "name": (name or "AI_Output").strip() or "AI_Output",
        "cells": normalized_cells,
        "formulas": normalized_formulas,
    }

    # Optional styling hints
    if conditional_formatting:
        payload["conditional_formatting"] = conditional_formatting
    if column_widths:
        payload["column_widths"] = column_widths
    if freeze_panes:
        payload["freeze_panes"] = freeze_panes

    return payload


def normalize_formula_targets(
    formulas: List[Dict[str, Any]],
    start_col: str = "F",
    start_row: int = 2
) -> List[Dict[str, Any]]:
    """
    Ensures each formula has a concrete 'address', assigning sequential cells if missing.
    This helps avoid overwriting the data region; defaults to F2 downward.
    """
    row = start_row
    out: List[Dict[str, Any]] = []
    for f in formulas or []:
        addr = f.get("address") or f.get("cell")
        if not addr:
            addr = f"{start_col}{row}"
            row += 1
        form = normalize_formula_text(f.get("formula", ""))
        if form:
            entry = {"address": addr, "formula": form}
            desc = f.get("description")
            if isinstance(desc, str) and desc:
                entry["description"] = desc
            out.append(entry)
    return out


def combine_operation_results(
    worksheets: List[Dict[str, Any]],
    message: str = "Excel operations prepared."
) -> Dict[str, Any]:
    """
    Combine multiple worksheet payloads into an ActionResponse-compatible dict:
    {
      "success": true,
      "message": "...",
      "worksheets": [...],
      "formulas": []  # kept empty to avoid duplicated top-level vs per-worksheet
    }
    """
    return {
        "success": True,
        "message": message,
        "worksheets": worksheets or [],
        "formulas": [],  # formulas live inside each worksheet payload
    }


# ==========================
# High-level helper routines
# ==========================

def make_simple_table_sheet(
    name: str,
    headers: List[str],
    rows: List[List[Any]],
    start_cell: str = "A1",
    autosize: bool = True,
    freeze_header: bool = True,
    notes: str = ""
) -> Dict[str, Any]:
    """
    Build a simple tabular worksheet payload by writing headers and data cells.
    This does not create a formal Excel Table object—frontend may choose to do so.

    Args:
      name: Worksheet name.
      headers: Column headers to place at the starting row.
      rows: Data rows aligned to headers length.
      start_cell: Top-left cell address where the header row is placed.
      autosize: If True, include column_widths hint to auto-fit (frontend can implement).
      freeze_header: If True, include freeze_panes at the row below header.
      notes: Optional notes at A1/A2.

    Returns:
      Worksheet payload with cells filled for headers and data.
    """
    # Extract start column letter and row from A1 style
    match = re.match(r"^([A-Za-z]+)([0-9]+)$", start_cell or "A1")
    start_col = match.group(1) if match else "A"
    start_row = int(match.group(2)) if match else 1

    def col_to_index(col_letters: str) -> int:
        acc = 0
        for ch in col_letters.upper():
            acc = acc * 26 + (ord(ch) - 64)
        return acc

    def index_to_col(idx: int) -> str:
        letters = ""
        while idx > 0:
            idx, r = divmod(idx - 1, 26)
            letters = chr(65 + r) + letters
        return letters

    start_idx = col_to_index(start_col)
    cells: List[Dict[str, Any]] = []

    # Header row
    for j, h in enumerate(headers or []):
        col_idx = start_idx + j
        addr = f"{index_to_col(col_idx)}{start_row}"
        cells.append({"address": addr, "value": h})

    # Data rows
    for i, row in enumerate(rows or [], start=1):
        for j, val in enumerate(row[: len(headers or row)]):
            col_idx = start_idx + j
            addr = f"{index_to_col(col_idx)}{start_row + i}"
            cells.append({"address": addr, "value": val})

    column_widths = None
    if autosize and headers:
        # Give a generic width hint; frontends can auto-fit or ignore
        column_widths = [(index_to_col(start_idx + j), max(10.0, min(28.0, float(len(str(h or '')) + 4)))) for j, h in enumerate(headers)]

    freeze_panes = None
    if freeze_header:
        freeze_panes = f"{start_col}{start_row + 1}"

    return build_worksheet_payload(
        name=name or "AI_Table",
        cells=cells,
        formulas=[],
        notes=notes,
        conditional_formatting=None,
        column_widths=column_widths,
        freeze_panes=freeze_panes,
    )


def make_formula_sheet(
    name: str,
    formulas: List[Dict[str, Any]],
    notes: str = "",
    start_col: str = "F",
    start_row: int = 2
) -> Dict[str, Any]:
    """
    Build a worksheet payload with one or more formulas, normalized for safe placement.
    Each formula item may be {"address":"H2","formula":"=SUM(A2:A100)","description":"..."}.
    If address is absent, the function assigns sequential addresses from start_col/start_row.
    """
    normalized = normalize_formula_targets(formulas, start_col=start_col, start_row=start_row)
    return build_worksheet_payload(
        name=name or "AI_Formulas",
        cells=[],
        formulas=normalized,
        notes=notes,
    )


def make_notes_sheet(
    name: str,
    title: str,
    lines: List[str],
    start_cell: str = "A1",
    notes: str = ""
) -> Dict[str, Any]:
    """
    Build a simple notes sheet with a title and multi-line content.
    """
    match = re.match(r"^([A-Za-z]+)([0-9]+)$", start_cell or "A1")
    col = match.group(1) if match else "A"
    row = int(match.group(2)) if match else 1

    cells: List[Dict[str, Any]] = [
        {"address": f"{col}{row}", "value": title or "Notes"},
    ]
    for i, line in enumerate(lines or [], start=1):
        cells.append({"address": f"{col}{row + i}", "value": line})

    return build_worksheet_payload(
        name=name or "AI_Notes",
        cells=cells,
        formulas=[],
        notes=notes,
    )
