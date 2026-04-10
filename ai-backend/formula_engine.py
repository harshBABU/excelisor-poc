from typing import Any, Dict, List, Tuple
import os
from llm_config import get_model

# Optional: integrate with your existing OpenAI client wrapper if present.
try:
    from openai import OpenAI
except Exception:
    OpenAI = None


def _col_letter(idx: int) -> str:
    idx += 1
    letters = ""
    while idx:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def infer_headers_and_data(table: List[List[Any]]) -> Tuple[List[str], List[List[Any]]]:
    if not table:
        return [], []
    first = table[0]
    looks_like_header = (
        len(first) > 0
        and all(isinstance(x, str) for x in first)
        and len(set(first)) == len(first)
    )
    if looks_like_header and len(table) > 1:
        headers = [c if c else f"Col{i+1}" for i, c in enumerate(first)]
        data = table[1:]
    else:
        max_len = max((len(r) for r in table), default=0)
        headers = [f"Col{i+1}" for i in range(max_len)]
        data = [r + [None] * (max_len - len(r)) for r in table]
    return headers, data


def suggest_basic_formulas(table: List[List[Any]]) -> Dict[str, Any]:
    headers, data = infer_headers_and_data(table)
    if not headers or not data:
        return {
            "worksheet_name": "AI_Formulas",
            "formulas": [],
            "notes": "No data available for formula suggestions."
        }

    rows = len(data)
    cols = len(headers)

    formulas: List[Dict[str, str]] = []
    numeric_col_indices: List[int] = []
    text_col_indices: List[int] = []

    for j in range(cols):
        col_values = [row[j] for row in data if row[j] is not None]
        is_numeric = False
        if col_values:
            is_numeric = all(isinstance(v, (int, float)) for v in col_values)
        if is_numeric:
            numeric_col_indices.append(j)
        else:
            text_col_indices.append(j)

    notes = []
    row_end = rows + 1  # data rows occupy 2..rows+1 if row1 is header

    # Simple numeric summary: SUM and AVERAGE for first numeric col
    if numeric_col_indices:
        n_idx = numeric_col_indices[0]
        n_letter = _col_letter(n_idx)
        n_range = f"{n_letter}2:{n_letter}{row_end}"
        formulas.append({"address": "F2", "formula": f"=SUM({n_range})"})
        formulas.append({"address": "F3", "formula": f"=AVERAGE({n_range})"})
        notes.append(f"SUM and AVERAGE for first numeric column ({headers[n_idx]}).")

    # Simple COUNTIF scaffold if we have a text column
    if text_col_indices:
        t_idx = text_col_indices[0]
        t_letter = _col_letter(t_idx)
        t_range = f"{t_letter}2:{t_letter}{row_end}"
        formulas.append({"address": "G1", "formula": "=\"EnterCategory\""})
        formulas.append({"address": "G2", "formula": f"=COUNTIF({t_range}, G1)"})
        notes.append(f"COUNTIF for first text column ({headers[t_idx]}) using G1 as category input.")

    return {
        "worksheet_name": "AI_Formulas",
        "formulas": formulas,
        "notes": " ".join(notes) if notes else "Starter formulas."
    }


def generate_ai_formulas(table: List[List[Any]], request_hint: str = "") -> Dict[str, Any]:
    """
    Fallback: ask AI for 2-4 safe formulas that operate on the selection.
    Constraints:
    - Use English Excel functions (SUM, AVERAGE, COUNTIF, SUMIF, MAX, MIN, etc.).
    - Assume headers in row 1, data from row 2..N.
    - Use letter ranges (e.g., A2:A100) for broad compatibility.
    - Provide absolute addresses to place the formulas (e.g., H2, H3, ...), avoid overwriting the selected range.
    """
    try:
        from llm_config import get_openai_client
        client = get_openai_client()
    except Exception:
        # If no AI available, return empty so caller can handle gracefully.
        return {
            "worksheet_name": "AI_Formulas",
            "formulas": [],
            "notes": "AI not configured; no additional formulas generated."
        }

    headers, data = infer_headers_and_data(table)
    rows = len(data)
    cols = len(headers)
    row_end = rows + 1 if rows > 0 else 2

    # Build a compact, bounded preview to keep prompts small
    def csv_preview(max_rows=6):
        out = []
        out.append(",".join(headers) if headers else "")
        for r in data[:max_rows]:
            out.append(",".join("" if v is None else str(v) for v in r[:cols]))
        return "\n".join(out)

    preview = csv_preview()

    system = (
        "You output only a small set of Excel formulas placed into target cells."
        " Use English Excel function names. Assume headers in row1, data starts at row2."
        " Use column letters (A, B, C...). Avoid volatile or locale-sensitive functions."
        " Prefer SUM, AVERAGE, COUNTIF, SUMIF, MAX, MIN, MEDIAN, STDEV where applicable."
        " Return 2-4 formulas max."
    )
    # Simple JSON-like plan the model should fill.
    # We won't do function-calling here to avoid coupling; we will parse minimally.
    user = (
        f"Selection preview (CSV):\n{preview}\n\n"
        f"Rows of data: {rows}, Columns: {cols}.\n"
        f"Place formulas in columns starting at H2 (H2, H3, H4...), do not overwrite A..{_col_letter(cols-1)}.\n"
        f"Request hint: {request_hint or '(none)'}\n\n"
        "Create 2-4 useful formulas with absolute cell addresses. "
        "Respond strictly as lines of 'ADDRESS = FORMULA' (e.g., 'H2 = =SUM(A2:A101)'). No extra text."
    )


    try:
        resp = client.chat.completions.create(
            model=get_model(),
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user}
            ],
            temperature=0.1,
        )
        text = (resp.choices[0].message.content or "").strip()
    except Exception as e:
        return {
            "worksheet_name": "AI_Formulas",
            "formulas": [],
            "notes": f"AI error: {e}"
        }

    formulas: List[Dict[str, str]] = []
    if text:
        for line in text.splitlines():
            line = line.strip()
            # Expect "H2 = =SUM(A2:A101)" or "H3= =AVERAGE(B2:B101)"
            if "=" in line:
                parts = line.split("=", 1)
                left = parts[0].strip()
                right = parts[1].strip()
                address = left.replace(" ", "")
                formula = right
                if not formula.startswith("="):
                    formula = "=" + formula
                # Simple validation: address must start with a letter
                if address and address[0].isalpha():
                    formulas.append({"address": address, "formula": formula})

    notes = "AI-generated formulas" if formulas else "AI returned no usable formulas."
    return {
        "worksheet_name": "AI_Formulas",
        "formulas": formulas[:4],
        "notes": notes
    }


def suggest_or_generate_formulas(table: List[List[Any]], request_hint: str = "") -> Dict[str, Any]:
    """
    Try rule-based first; if nothing or very little is produced, fallback to AI.
    """
    plan = suggest_basic_formulas(table)
    if len(plan.get("formulas", [])) >= 2:
        return plan
    ai_plan = generate_ai_formulas(table, request_hint=request_hint)
    if ai_plan.get("formulas"):
        return ai_plan
    return plan
