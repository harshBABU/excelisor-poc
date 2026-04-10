# aibackend/charting.py

from __future__ import annotations

import logging
import warnings
from typing import List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)
if not logger.handlers:
    if logging.getLogger().handlers:
        logger.propagate = True
    else:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

CHART_LINE = "Line"
CHART_COLUMN_CLUSTERED = "ColumnClustered"
CHART_BAR_CLUSTERED = "BarClustered"
CHART_SCATTER = "XYScatter"
CHART_PIE = "Pie"

def _is_datetime_like(series: pd.Series, min_ratio: float = 0.5) -> bool:
    try:
        s = series
        total = int(s.notna().sum())
        logger.debug("_is_datetime_like: starting check; total_non_null=%d, min_ratio=%.2f", total, min_ratio)
        if total == 0:
            logger.debug("_is_datetime_like: no non-null values")
            return False

        COMMON_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y")
        for fmt in COMMON_FORMATS:
            parsed = pd.to_datetime(s, errors="coerce", format=fmt)
            valid = int(parsed.notna().sum())
            ratio = valid / total if total else 0
            logger.debug("_is_datetime_like: explicit format %s valid=%d ratio=%.2f", fmt, valid, ratio)
            if valid >= 3 and ratio >= min_ratio:
                logger.debug("_is_datetime_like: matched explicit format %s", fmt)
                return True

        try:
            iso = pd.to_datetime(s, errors="coerce", format="ISO8601")
            valid_iso = int(iso.notna().sum())
            ratio_iso = valid_iso / total
            logger.debug("_is_datetime_like: ISO8601 parsed valid=%d ratio=%.2f", valid_iso, ratio_iso)
            if valid_iso >= 3 and ratio_iso >= min_ratio:
                logger.debug("_is_datetime_like: ISO8601 match")
                return True
        except Exception as e:
            logger.debug("_is_datetime_like: ISO8601 parse error: %s", e)

        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            strict = pd.to_datetime(s, errors="coerce")
        valid_strict = int(strict.notna().sum())
        ratio_strict = valid_strict / total
        logger.debug("_is_datetime_like: strict parse valid=%d ratio=%.2f", valid_strict, ratio_strict)
        if valid_strict >= 3 and ratio_strict >= min_ratio:
            logger.debug("_is_datetime_like: strict parse match")
            return True

        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            mixed = pd.to_datetime(s, errors="coerce", format="mixed")
        valid_mixed = int(mixed.notna().sum())
        ratio_mixed = valid_mixed / total
        logger.debug("_is_datetime_like: mixed parse valid=%d ratio=%.2f", valid_mixed, ratio_mixed)
        if valid_mixed >= 3 and ratio_mixed >= min_ratio:
            logger.debug("_is_datetime_like: mixed parse match")
            return True

        logger.debug("_is_datetime_like: no datetime-like parse matched")
        return False
    except Exception as e:
        logger.exception("_is_datetime_like: unexpected error")
        return False

def _find_datetime_col(df: pd.DataFrame, min_ratio: float = 0.5) -> Optional[str]:
    logger.debug("_find_datetime_col: checking %d columns", len(df.columns))
    for c in df.columns:
        logger.debug("_find_datetime_col: checking column '%s' for datetime-like values", c)
        try:
            if _is_datetime_like(df[c], min_ratio=min_ratio):
                logger.info("_find_datetime_col: datetime-like detected: %s", c)
                return c
        except Exception as e:
            logger.warning("_find_datetime_col: error checking column '%s': %s", c, e)
    logger.debug("_find_datetime_col: no datetime-like column found")
    return None

def _list_numeric_cols(df: pd.DataFrame) -> List[str]:
    try:
        cols = df.select_dtypes(include=["number"]).columns.tolist()
        logger.debug("_list_numeric_cols: selected numeric columns=%s", cols)
        return cols
    except Exception as e:
        logger.warning("_list_numeric_cols: failed to list numeric columns: %s", e)
        return []

def _list_categorical_cols(df: pd.DataFrame, max_cardinality: int = 50) -> List[str]:
    cats: List[str] = []
    for c in df.columns:
        try:
            s = df[c]
            if s.dtype == object or pd.api.types.is_string_dtype(s):
                nunique = int(s.nunique(dropna=True))
                logger.debug("_list_categorical_cols: column=%s dtype=%s nunique=%d", c, s.dtype, nunique)
                if 0 < nunique <= max_cardinality:
                    cats.append(c)
                else:
                    logger.debug("_list_categorical_cols: excluded %s by cardinality=%d", c, nunique)
        except Exception as e:
            logger.warning("_list_categorical_cols: column '%s' error: %s", c, e)
            continue
    logger.debug("_list_categorical_cols: selected categorical columns=%s", cats)
    return cats

def suggest_chart(df: pd.DataFrame) -> Tuple[str, str]:
    logger.debug("suggest_chart: Starting chart suggestion.")
    try:
        if df is None:
            logger.warning("suggest_chart: DataFrame is None")
            return CHART_COLUMN_CLUSTERED, "AI-Suggested Chart"
        if df.empty or len(df.columns) == 0:
            logger.warning("suggest_chart: Empty DataFrame or no columns")
            return CHART_COLUMN_CLUSTERED, "AI-Suggested Chart"

        logger.debug("suggest_chart: DataFrame shape=%s columns=%s", df.shape, list(df.columns))
        dt_col = _find_datetime_col(df, min_ratio=0.5)
        num_cols = _list_numeric_cols(df)
        cat_cols = _list_categorical_cols(df, max_cardinality=50)

        logger.info("suggest_chart: Results dt=%s num_cols=%s cat_cols=%s", dt_col, num_cols, cat_cols)

        if dt_col and len(num_cols) >= 1:
            chart_title = f"Trend over time ({dt_col})"
            logger.info("suggest_chart: selected %s title=%s", CHART_LINE, chart_title)
            return CHART_LINE, chart_title

        if len(cat_cols) >= 1 and len(num_cols) >= 1:
            chart_title = f"{cat_cols[0]} by {num_cols[0]}"
            logger.info("suggest_chart: selected %s title=%s", CHART_COLUMN_CLUSTERED, chart_title)
            return CHART_COLUMN_CLUSTERED, chart_title

        if len(num_cols) >= 2:
            chart_title = f"{num_cols[0]} vs {num_cols[1]}"
            logger.info("suggest_chart: selected %s title=%s", CHART_COLUMN_CLUSTERED, chart_title)
            return CHART_COLUMN_CLUSTERED, chart_title

        if len(num_cols) == 1:
            chart_title = f"{num_cols[0]}"
            logger.info("suggest_chart: selected %s title=%s", CHART_COLUMN_CLUSTERED, chart_title)
            return CHART_COLUMN_CLUSTERED, chart_title

        logger.warning("suggest_chart: fallback chart selected")
        return CHART_COLUMN_CLUSTERED, "AI-Suggested Chart"

    except Exception as e:
        logger.exception("suggest_chart: unexpected error")
        return CHART_COLUMN_CLUSTERED, "AI-Suggested Chart"
