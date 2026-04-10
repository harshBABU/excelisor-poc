# aibackend/table_utils.py

from __future__ import annotations

import logging
from typing import Any, Dict, List, Sequence

import numpy as np
import pandas as pd


logger = logging.getLogger(__name__)
if not logger.handlers:
    # Keep logging lightweight; main app can override level/handlers
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


def to_python_scalars(obj: Any) -> Any:
    """
    Recursively convert numpy/pandas scalar types to native Python types
    to avoid JSON serialization issues in FastAPI responses.

    Examples:
      - np.int64(5) -> 5
      - np.float64(3.14) -> 3.14
      - dicts/lists are traversed recursively
    """
    # numpy scalar
    if isinstance(obj, (np.generic,)):
        try:
            return obj.item()
        except Exception:
            # Fallback to Python conversion if item() not available
            return obj.tolist() if hasattr(obj, "tolist") else obj

    # pandas NA types handled as None
    try:
        # pandas._libs.missing.NAType might not always be available; use check via pandas.isna
        if obj is pd.NA:  # type: ignore[attr-defined]
            return None
    except Exception:
        pass

    # dict recursion
    if isinstance(obj, dict):
        return {k: to_python_scalars(v) for k, v in obj.items()}

    # list/tuple recursion
    if isinstance(obj, (list, tuple)):
        return [to_python_scalars(v) for v in obj]

    return obj


def _looks_like_header(row: Sequence[Any]) -> bool:
    """
    Heuristic: a row looks like headers if all values are strings and unique,
    and there is at least one element.
    """
    if not row or not isinstance(row, (list, tuple)):
        return False
    if not all(isinstance(x, str) for x in row):
        return False
    # uniqueness helps avoid mistaking data rows for headers
    return len(set(row)) == len(row)


def _normalize_rows(table_2d: List[List[Any]]) -> List[List[Any]]:
    """
    Normalize ragged rows to the same length by padding with None.
    """
    max_len = max((len(r) for r in table_2d), default=0)
    if max_len == 0:
        return []
    normalized: List[List[Any]] = []
    for r in table_2d:
        # Ensure each row is a list
        row_list = list(r) if isinstance(r, (list, tuple)) else [r]
        if len(row_list) < max_len:
            row_list = row_list + [None] * (max_len - len(row_list))
        normalized.append(row_list)
    return normalized


def table_to_dataframe(table_2d: List[List[Any]]) -> pd.DataFrame:
    """
    Convert a 2D list (Excel selection values) into a pandas DataFrame.

    Behavior:
      - If the first row looks like headers (all strings, unique) AND there is at least one more row,
        use that row as DataFrame columns and the remaining rows as data.
      - Otherwise, create generic headers: Col1..ColN and include every row as data.
      - Ragged rows are padded with None.

    Args:
      table_2d: A 2D list as returned by Excel JS API's range.values (list of rows).

    Returns:
      pandas.DataFrame with inferred or generated headers.
    """
    try:
        if not table_2d or not isinstance(table_2d, list):
            logger.debug("table_to_dataframe: empty or invalid table input; returning empty DataFrame.")
            return pd.DataFrame()

        # Ensure all rows are lists and normalize ragged rows
        normalized = _normalize_rows(
            [list(r) if isinstance(r, (list, tuple)) else [r] for r in table_2d]
        )

        if not normalized:
            return pd.DataFrame()

        first = normalized[0]
        has_data_beyond_header = len(normalized) > 1
        if _looks_like_header(first) and has_data_beyond_header:
            df = pd.DataFrame(normalized[1:], columns=first)
            logger.debug(
                "table_to_dataframe: using first row as headers; rows=%d, cols=%d",
                df.shape[0], df.shape[1]
            )
            return df

        # Fallback: generic headers
        max_len = len(first)
        cols = [f"Col{i+1}" for i in range(max_len)]
        df = pd.DataFrame(normalized, columns=cols)
        logger.debug(
            "table_to_dataframe: generated generic headers; rows=%d, cols=%d",
            df.shape[0], df.shape[1]
        )
        return df

    except Exception as e:
        # Never crash caller; return empty DataFrame with a log
        logger.error(f"table_to_dataframe: failed to convert table -> DataFrame: {e}")
        return pd.DataFrame()
