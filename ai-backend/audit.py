# aibackend/audit.py

from __future__ import annotations

import logging
from typing import List

import json
import os

import numpy as np
import pandas as pd


logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
# Keep INFO by default; callers can override to DEBUG for more detail
logger.setLevel(logging.INFO)


import json
from pydantic import ValidationError
from models import ActionResponse, WorksheetPayload, CellWrite, ConditionalFormatHint

logger = logging.getLogger(__name__)

def audit_dataframe(df: pd.DataFrame) -> dict:
    """
    Perform a dynamically determined audit on the DataFrame using an LLM.
    Returns a dict that fits `ActionResponse` so the frontend can draw the dynamic worksheet.
    """
    from llm_config import get_openai_client, get_model
    client = get_openai_client()
    
    logger.info(f"[AUDIT] Starting dynamic audit on shape: {df.shape}")
    if df is None or df.empty:
        return {"success": False, "message": "No data selected for audit.", "worksheets": []}

    # 1. Profile and standardize columns
    safe_cols = {c: str(c).replace(" ", "_") for c in df.columns}
    # map back for later
    orig_cols = {v: k for k, v in safe_cols.items()}
    df_safe = df.rename(columns=safe_cols)
    
    dtypes_info = {str(col): str(dtype) for col, dtype in df_safe.dtypes.items()}
    sample_records = df_safe.head(3).to_dict(orient="records")

    profile = {
        "columns": list(df_safe.columns),
        "dtypes": dtypes_info,
        "sample": sample_records,
        "total_rows": len(df_safe)
    }

    # 2. Ask LLM to determine metrics
    # We ask for valid pandas queries (boolean masks for df_safe) that identify BAD rows.
    prompt = f"""You are a brilliant Data Quality Analyst. 
Given the following data profile, determine 3 to 6 highly specific data quality checks that make sense for this exact dataset.

Profile:
{json.dumps(profile, indent=2)}

Stipulations:
1. Output ONLY a valid JSON object containing a "checks" array. 
2. Use EXACTLY this format:
{{
  "checks": [
    {{
      "name": "Negative Quantities in Orders",
      "description": "Quantities should never be negative.",
      "pandas_query": "Quantity < 0"
    }}
  ]
}}
3. 'pandas_query' must be a valid string for `df.query()`. It must resolve to a boolean mask where True = FAULTY/VIOLATING row.
4. Only use the column names provided in the profile (they have been safely underscore-formatted).
"""
    try:
        resp = client.chat.completions.create(
            model=get_model(),
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"} if "gpt-4" in get_model().lower() or "gpt-5" in get_model().lower() else None
        )
        text = resp.choices[0].message.content.strip()
        logger.debug(f"[AUDIT] LLM raw response: {text}")
        if text.startswith("```"): text = text.split("```")[1]
        if text.startswith("json"): text = text.replace("json", "", 1).strip()
        
        parsed = json.loads(text)
        checks = []
        if isinstance(parsed, dict):
            if "checks" in parsed and isinstance(parsed["checks"], list):
                checks = parsed["checks"]
            else:
                # search for any generic list
                for k, v in parsed.items():
                    if isinstance(v, list):
                        checks = v
                        break
        elif isinstance(parsed, list):
            checks = parsed
            
        if not checks or not isinstance(checks, list):
            raise ValueError(f"LLM did not return a JSON array. Parsed: {parsed}")
    except Exception as e:
        logger.error(f"[AUDIT] LLM Metric Determination Failed: {e}")
        return {"success": False, "message": "Failed to dynamically determine audit rules.", "worksheets": []}

    # 3. Calculate Metrics securely
    results = []
    total_faults = 0
    for ck in checks:
        query_str = ck.get("pandas_query", "")
        if not query_str: continue
        
        try:
            faulty_df = df_safe.query(query_str)
            fault_count = len(faulty_df)
            total_faults += fault_count
            
            # optionally capture sample faulty indices
            sample_faulty = faulty_df.head(2).to_dict(orient="records")
            
            results.append({
                "name": ck.get("name", "Unknown Check"),
                "description": ck.get("description", ""),
                "query": query_str,
                "count": fault_count,
                "sample": sample_faulty
            })
        except Exception as e:
            logger.warning(f"[AUDIT] Check '{ck.get('name')}' failed execution on query '{query_str}': {e}")
            results.append({
                "name": ck.get("name", ""),
                "description": ck.get("description", ""),
                "count": "Error",
                "error": str(e)
            })

    # 4. Plan the summary Worksheet Payload
    cells = []
    
    # Title
    cells.append(CellWrite(address="A1", value="AI Data Quality Audit & Visualization"))
    
    # Headers
    cells.append(CellWrite(address="A3", value="Metric / Check Name"))
    cells.append(CellWrite(address="B3", value="Violations Found"))
    cells.append(CellWrite(address="C3", value="Description"))
    
    row_idx = 4
    for r in results:
        cells.append(CellWrite(address=f"A{row_idx}", value=r.get("name")))
        cells.append(CellWrite(address=f"B{row_idx}", value=r.get("count")))
        cells.append(CellWrite(address=f"C{row_idx}", value=r.get("description")))
        row_idx += 1
        
    last_row = row_idx - 1

    # Formatting and Visualization Layout
    # Apply a DataBar conditional format to Visually highlight the degree of exceptions/errors
    cf = []
    if last_row >= 4:
        cf.append(ConditionalFormatHint(
            range=f"B4:B{last_row}",
            type="dataBar"
        ))
        
    # Append some sample bad data
    row_idx += 2
    cells.append(CellWrite(address=f"A{row_idx}", value="Sample of Violating Data Rows"))
    row_idx += 1
    
    # Build column headers for the sample data chunk
    col_letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for i, col_name in enumerate(orig_cols.values()):
        if i < len(col_letters):
            cells.append(CellWrite(address=f"{col_letters[i]}{row_idx}", value=col_name))
    row_idx += 1
    
    samples_written = 0
    for r in results:
        if r.get("count") != "Error" and int(r.get("count")) > 0:
            for bad_row in r.get("sample", []):
                # Write back translating to original column keys
                for i, safe_k in enumerate(df_safe.columns):
                    if i < len(col_letters):
                        val = bad_row.get(safe_k, "")
                        # coerce nan/nat to empty string
                        if pd.isna(val): val = ""
                        # Convert bools to strings or int/floats as is
                        if type(val) not in (int, float, str, bool): val = str(val)
                        cells.append(CellWrite(address=f"{col_letters[i]}{row_idx}", value=val))
                row_idx += 1
                samples_written += 1
                if samples_written >= 10: break
        if samples_written >= 10: break

    worksheet = WorksheetPayload(
        name="AI_Audit_Scorecard",
        cells=cells,
        conditional_formatting=cf
    )

    action_res = {
        "success": True,
        "message": f"Audit complete. Found {total_faults} total violations across {len(results)} metrics. See 'AI_Audit_Scorecard' for details and visualizations.",
        "worksheets": [worksheet.model_dump()]
    }

    return action_res