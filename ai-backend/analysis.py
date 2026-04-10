from __future__ import annotations

import json
import logging
import os
from typing import Optional, Dict, Any, List

import pandas as pd
from fastapi import HTTPException

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

from llm_config import get_model


# =========================
# Logging
# =========================
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


# =========================
# OpenAI Client
# =========================
def _get_client() -> Optional["OpenAI"]:
    try:
        from llm_config import get_openai_client
        return get_openai_client()
    except Exception as e:
        logger.error(f"OpenAI client init error: {e}")
        return None


# =========================
# Allowed Config
# =========================
ALLOWED_METRICS = [
    "sum", "avg", "count", "min", "max",
    "median", "std", "variance"
]

ALLOWED_STEPS = {
    "aggregate", "derive", "sort",
    "top_n", "bottom_n", "rank",
    "growth", "share_of_total"
}

# Shown to LLM in planner prompt so it knows the exact schema
PLAN_FORMAT_EXAMPLE = """
VALID STEP TYPES — use EXACTLY these string values for "step":

  {"step": "derive", "new_column": "TOTAL_SALES", "formula": "RETAIL_SALES + WAREHOUSE_SALES"}
  {"step": "aggregate", "group_by": ["MONTH"], "aggregations": [{"column": "TOTAL_SALES", "metric": "sum"}]}
  {"step": "sort", "column": "TOTAL_SALES", "order": "desc"}
  {"step": "top_n", "column": "TOTAL_SALES", "top_n": 1}
  {"step": "bottom_n", "column": "TOTAL_SALES", "bottom_n": 5}
  {"step": "rank", "column": "TOTAL_SALES"}
  {"step": "growth", "column": "TOTAL_SALES"}
  {"step": "share_of_total", "column": "TOTAL_SALES"}

NEVER use integers (1, 2, 3...) for "step".
NEVER invent new step types.
ONLY use: derive, aggregate, sort, top_n, bottom_n, rank, growth, share_of_total.
"""


# =========================
# Data Profiling
# =========================
def _profile_dataframe(df: pd.DataFrame) -> Dict[str, Any]:
    logger.info("=" * 60)
    logger.info("[STEP 1] PROFILING DATAFRAME")
    logger.info(f"  Shape        : {df.shape[0]} rows x {df.shape[1]} columns")
    logger.info(f"  Columns      : {df.columns.tolist()}")
    logger.info(f"  Dtypes       :\n{df.dtypes.to_string()}")
    logger.info(f"  Null counts  :\n{df.isnull().sum().to_string()}")
    # logger.info(f"  Sample (3)   :\n{df.head(3).to_string()}")  # verbose

    profile = {
        "columns": df.columns.tolist(),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        "sample": df.head(3).to_dict(orient="records"),
        "unique_counts": df.nunique().to_dict(),
    }

    time_cols = [
        col for col in df.columns
        if any(k in col.lower() for k in ["date", "time", "year", "month"])
    ]
    profile["time_columns"] = time_cols

    if "YEAR" in df.columns and "MONTH" in df.columns:
        profile["time_combination"] = ["YEAR", "MONTH"]

    # FIX: safe column names (spaces → underscores) for eval safety
    profile["safe_columns"] = {col: col.replace(" ", "_") for col in df.columns}

    logger.info(f"  Time columns : {time_cols}")
    logger.info(f"  Safe columns : {profile['safe_columns']}")
    logger.info("[STEP 1] PROFILING COMPLETE")
    return profile


# =========================
# Semantic Understanding
# =========================
def _infer_semantics(profile: Dict[str, Any]) -> Dict[str, Any]:
    logger.info("=" * 60)
    logger.info("[STEP 2] INFERRING SEMANTICS via LLM")

    client = _get_client()
    if not client:
        raise HTTPException(400, "Missing OPENAI_API_KEY")

    prompt = f"""
Classify each column into:
- measure
- dimension
- time
- id

Return ONLY JSON.

Profile:
{json.dumps(profile)}
"""

    resp = client.responses.create(
        model=get_model(),
        input=prompt
    )

    text = resp.output_text.strip()
    logger.info(f"  LLM raw semantics response:\n{text}")

    try:
        semantics = json.loads(text)
        logger.info(f"  Parsed semantics: {semantics}")
        return semantics
    except Exception as e:
        logger.warning(f"  Failed to parse semantics: {e}. Returning empty.")
        return {}


# =========================
# Rewrite Plan Column Names
# =========================
def _rewrite_plan_columns(plan, col_rename):
    """Replace spaced column names with underscore-safe versions in the plan."""
    import copy
    plan = copy.deepcopy(plan)
    for step in plan:
        for key in ["column", "new_column"]:
            if key in step:
                val = step[key]
                if isinstance(val, str):
                    step[key] = col_rename.get(val, val)
        if "group_by" in step:
            val = step["group_by"]
            if isinstance(val, list):
                step["group_by"] = [col_rename.get(c, c) for c in val]
            elif isinstance(val, str):
                step["group_by"] = col_rename.get(val, val)
        if "aggregations" in step:
            for agg in step["aggregations"]:
                agg["column"] = col_rename.get(agg["column"], agg["column"])
        if "formula" in step:
            for old, new in col_rename.items():
                step["formula"] = step["formula"].replace(old, new)
    return plan


# =========================
# Planner
# =========================
def _plan(question, profile, semantics, error=None):
    logger.info("=" * 60)
    logger.info("[STEP 3] PLANNING via LLM")
    logger.info(f"  Question : {question}")
    logger.info(f"  Error    : {error}")

    client = _get_client()
    if not client:
        raise HTTPException(400, "Missing OPENAI_API_KEY")

    prompt = f"""
You are a data planner. Return a JSON list of execution steps to answer the question.

{PLAN_FORMAT_EXAMPLE}

STRICT RULES:
- "step" MUST be one of the valid string types above — never an integer
- derive MUST include "formula" and "new_column"
- aggregate MUST include "group_by" and "aggregations"
- Return ONLY a JSON list, no explanation, no markdown
- Use safe_columns names (underscores, not spaces) in all formulas and column references
- Only use top_n or bottom_n as the FINAL step if the question EXPLICITLY asks for "top N", "best N", "worst N", or "bottom N" results
- For open-ended questions like "give me insights", "analyze the data", "what are the trends", "interesting patterns" — do NOT use top_n or bottom_n — return ALL aggregated rows so the responder has full context to generate rich insights

Profile:
{json.dumps(profile)}

Semantics:
{json.dumps(semantics)}

Error from previous attempt (fix this):
{error}

Question:
{question}
"""

    resp = client.responses.create(
        model=get_model(),
        input=prompt
    )

    text = resp.output_text.strip()
    logger.info(f"  LLM raw plan response:\n{text}")

    # clean markdown
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]

    try:
        plan = json.loads(text)
        logger.info(f"  Parsed plan ({len(plan)} steps): {json.dumps(plan, indent=2)}")
        return plan
    except Exception:
        logger.error(f"  Planner JSON parse failed. Raw text: {text}")
        return []


# =========================
# Sanitize Plan (CRITICAL)
# =========================
def _sanitize_plan(plan):
    logger.info("=" * 60)
    logger.info("[STEP 4] SANITIZING PLAN")
    cleaned = []

    for i, step in enumerate(plan):
        if not isinstance(step, dict):
            logger.warning(f"  Step {i} skipped: not a dict → {step}")
            continue
        if "step" not in step:
            logger.warning(f"  Step {i} skipped: missing 'step' key → {step}")
            continue
        # FIX 1: Reject steps where "step" is not a valid string type
        if not isinstance(step["step"], str):
            logger.warning(f"  Step {i} skipped: 'step' is not a string (got {type(step['step']).__name__}: {step['step']!r}) → {step}")
            continue
        # FIX 1: Reject unknown step types
        if step["step"] not in ALLOWED_STEPS:
            logger.warning(f"  Step {i} skipped: unknown step type '{step['step']}' — not in ALLOWED_STEPS → {step}")
            continue
        if step["step"] == "derive" and "formula" not in step:
            logger.warning(f"  Step {i} skipped: derive missing 'formula' → {step}")
            continue
        logger.info(f"  Step {i} accepted: {step['step']}")
        cleaned.append(step)

    logger.info(f"  Sanitized plan: {len(cleaned)}/{len(plan)} steps kept")
    return cleaned


# =========================
# Default Plan (Fallback)
# =========================
def _default_plan(df):
    logger.warning("[FALLBACK] Using default plan — LLM plan was empty")
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    cat_cols = df.select_dtypes(include=["object"]).columns.tolist()

    if not numeric_cols:
        logger.warning("  No numeric columns found — returning empty plan")
        return []

    plan = [
        {
            "step": "aggregate",
            "group_by": cat_cols[:1],
            "aggregations": [
                {"column": numeric_cols[0], "metric": "sum"}
            ]
        },
        {
            "step": "top_n",
            "column": numeric_cols[0],
            "top_n": 5
        }
    ]
    logger.info(f"  Default plan: {json.dumps(plan, indent=2)}")
    return plan


# =========================
# LLM Intent Classification
# =========================
def _classify_intent(question: str, profile: Dict[str, Any]) -> Dict[str, Any]:
    """Ask the LLM whether this question needs a full grand-total aggregation."""
    logger.info("=" * 60)
    logger.info("[STEP 2b] CLASSIFYING INTENT via LLM")

    client = _get_client()
    if not client:
        raise HTTPException(400, "Missing OPENAI_API_KEY")

    prompt = f"""
You are a data intent classifier.

Given a question and a data profile, decide:
1. needs_full_aggregation: true if the question asks for a grand total, overall sum, or combined value across ALL rows. false if it asks for a breakdown, top N, trend, per-group values, or anything other than a single grand total.
2. aggregation_columns: list of column names from safe_columns that should be summed. Only populate if needs_full_aggregation is true.
3. reason: one sentence explaining your decision.

Return ONLY JSON in this exact format:
{{
  "needs_full_aggregation": true or false,
  "aggregation_columns": ["col1", "col2"],
  "reason": "..."
}}

Profile:
{json.dumps(profile)}

Question:
{question}
"""

    resp = client.responses.create(
        model=get_model(),
        input=prompt
    )

    text = resp.output_text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]

    logger.info(f"  LLM raw intent response:\n{text}")

    try:
        intent = json.loads(text)
        logger.info(f"  needs_full_aggregation : {intent.get('needs_full_aggregation')}")
        logger.info(f"  aggregation_columns    : {intent.get('aggregation_columns')}")
        logger.info(f"  reason                 : {intent.get('reason')}")
        return intent
    except Exception as e:
        logger.warning(f"  Intent parse failed: {e} — defaulting to no forced aggregation")
        return {"needs_full_aggregation": False, "aggregation_columns": [], "reason": "parse failed"}


def _force_aggregate_if_needed(intent: Dict[str, Any], df: pd.DataFrame) -> pd.DataFrame:
    """Use LLM intent decision to force a full sum aggregation if needed."""
    if not intent.get("needs_full_aggregation", False):
        logger.info("[STEP 5b] LLM intent: no forced aggregation needed — skipping")
        return df

    cols = intent.get("aggregation_columns", [])

    # Fall back to all numeric cols if LLM didn't specify any
    if not cols:
        cols = df.select_dtypes(include=["number"]).columns.tolist()
        logger.warning(f"  No aggregation_columns specified by LLM — using all numeric: {cols}")

    # Keep only columns that actually exist
    valid_cols = [c for c in cols if c in df.columns]
    missing = [c for c in cols if c not in df.columns]
    if missing:
        logger.warning(f"  Columns not found in df (skipped): {missing}")

    if not valid_cols:
        logger.warning("  No valid columns to aggregate — returning df as-is")
        return df

    logger.info("=" * 60)
    logger.info("[STEP 5b] FORCING FULL AGGREGATION per LLM intent")
    logger.info(f"  Summing columns: {valid_cols}")

    result = df[valid_cols].sum().to_frame(name="total").T
    # logger.info(f"  Aggregated result:\n{result.to_string()}")  # verbose
    logger.info("[STEP 5b] FORCED AGGREGATION COMPLETE")
    return result


# =========================
# Executor
# =========================
def _execute(plan, df):
    logger.info("=" * 60)
    logger.info("[STEP 5] EXECUTING PLAN")
    logger.info(f"  Input DataFrame shape: {df.shape}")

    current_df = df.copy()

    # FIX: Rename spaced column names to underscore-safe versions
    col_rename = {col: col.replace(" ", "_") for col in current_df.columns}
    col_restore = {v: k for k, v in col_rename.items()}
    renamed_cols = {k: v for k, v in col_rename.items() if k != v}
    if renamed_cols:
        logger.info(f"  Renaming columns for eval safety: {renamed_cols}")
        current_df.rename(columns=col_rename, inplace=True)

    # FIX: Rewrite plan to use safe column names
    plan = _rewrite_plan_columns(plan, col_rename)
    logger.info(f"  Plan after column rewrite: {json.dumps(plan, indent=2)}")

    # Create time column if YEAR+MONTH present
    if "YEAR" in current_df.columns and "MONTH" in current_df.columns:
        current_df["__time__"] = pd.to_datetime(
            current_df["YEAR"].astype(str) + "-" +
            current_df["MONTH"].astype(str) + "-01"
        )
        logger.info("  Created __time__ column from YEAR + MONTH")

    for i, step in enumerate(plan):
        step_name = step["step"]
        logger.info(f"  --- Step {i+1}/{len(plan)}: {step_name} ---")
        logger.info(f"      Config : {step}")
        logger.info(f"      DF shape before: {current_df.shape}")

        if step_name == "aggregate":
            agg_dict = {
                agg["column"]: {
                    "sum": "sum",
                    "avg": "mean",
                    "count": "count",
                    "min": "min",
                    "max": "max",
                    "median": "median",
                    "std": "std",
                    "variance": "var"
                }[agg["metric"]]
                for agg in step.get("aggregations", [])
            }
            group_by = step.get("group_by", [])
            logger.info(f"      group_by={group_by}, agg_dict={agg_dict}")

            # ── AGGREGATION HAPPENS HERE IN PYTHON (pandas groupby) ──
            current_df = current_df.groupby(group_by).agg(agg_dict).reset_index()
            # logger.info(f"      [PYTHON AGG] Result:\n{current_df.to_string()}")  # verbose

        elif step_name == "derive":
            formula = step["formula"]
            new_col = step["new_column"]
            logger.info(f"      [PYTHON DERIVE] Formula: {formula} → column: {new_col}")
            try:
                current_df[new_col] = current_df.eval(formula)
                # logger.info(f"      Derived column '{new_col}' sample: {current_df[new_col].head().tolist()}")  # verbose
            except Exception as e:
                logger.error(f"      Derive FAILED: {e}")
                raise Exception(f"Derive failed: {e}")

        elif step_name == "sort":
            col = step["column"]
            order = step.get("order", "desc")
            logger.info(f"      [PYTHON SORT] column={col}, order={order}")
            current_df = current_df.sort_values(
                by=col, ascending=(order == "asc")
            )

        elif step_name == "top_n":
            col = step["column"]
            n = step.get("top_n", 5)
            logger.info(f"      [PYTHON TOP_N] column={col}, n={n}")
            current_df = current_df.sort_values(by=col, ascending=False).head(n)

        elif step_name == "growth":
            col = step["column"]
            logger.info(f"      [PYTHON GROWTH] column={col}")
            if "__time__" in current_df.columns:
                current_df = current_df.sort_values("__time__")
            current_df["growth"] = current_df[col].pct_change()
            # logger.info(f"      Growth column sample: {current_df['growth'].head().tolist()}")  # verbose

        logger.info(f"      DF shape after : {current_df.shape}")

    # Restore original column names before returning
    restore_map = {v: k for k, v in col_rename.items() if k != v}
    if restore_map:
        current_df.rename(columns=restore_map, inplace=True)
        logger.info(f"  Restored column names: {restore_map}")

    # logger.info(f"  Final DataFrame ({current_df.shape}):\n{current_df.to_string()}")  # verbose
    logger.info("[STEP 5] EXECUTION COMPLETE")
    return current_df


# =========================
# Responder
# =========================
def _respond(question, df, semantics: Dict[str, Any] = None):
    logger.info("=" * 60)
    logger.info("[STEP 6] GENERATING RESPONSE via LLM")
    logger.info(f"  Result DataFrame shape: {df.shape}")

    # FIX: Never use describe() — always send actual values
    if len(df) <= 100:
        preview = df.to_csv(index=False)
        logger.info(f"  Sending full result ({len(df)} rows) to LLM")
    else:
        # FIX 2: Only sum measure columns — exclude time/id/dimension cols like YEAR, MONTH, ITEM CODE
        if semantics:
            measure_cols = [col for col, sem in semantics.items() if sem == "measure"]
            safe_measure_cols = [col.replace(" ", "_") for col in measure_cols]
            valid_cols = [c for c in safe_measure_cols if c in df.columns]
            logger.info(f"  Measure columns for fallback sum: {valid_cols}")
        else:
            valid_cols = df.select_dtypes(include=["number"]).columns.tolist()
            logger.warning("  No semantics provided — falling back to all numeric cols for sum")

        if not valid_cols:
            valid_cols = df.select_dtypes(include=["number"]).columns.tolist()

        summary = df[valid_cols].sum().to_frame(name="total").T
        preview = summary.to_csv(index=False)
        logger.warning(f"  DataFrame too large ({len(df)} rows) — sending measure sums to LLM")
        # logger.info(f"  Column sums:\n{summary.to_string()}")  # verbose

    # logger.info(f"  Preview sent to LLM:\n{preview}")  # verbose

    client = _get_client()
    if not client:
        raise HTTPException(400, "Missing OPENAI_API_KEY")

    prompt = f"""
You are a data analyst assistant. Answer the question using two sources:
1. DATA: the CSV below — use this for any numerical or data-driven part of the question.
2. YOUR KNOWLEDGE: use your general world knowledge for anything the data cannot answer (e.g. festivals, holidays, events, context).

Always clearly separate what comes from data vs what comes from your general knowledge.

Data:
{preview}

Question:
{question}
"""

    resp = client.responses.create(
        model=get_model(),
        input=prompt
    )

    answer = resp.output_text.strip()
    logger.info(f"  LLM answer: {answer}")
    logger.info("[STEP 6] RESPONSE COMPLETE")
    return answer


# =========================
# Agent
# =========================
def analyze_dataframe_agent(question: str, df: pd.DataFrame) -> str:
    logger.info("=" * 60)
    logger.info("[AGENT START]")
    logger.info(f"  Question : {question}")
    logger.info(f"  DF shape : {df.shape}")

    if not question.strip():
        raise HTTPException(400, "Question required")

    profile = _profile_dataframe(df)
    semantics = _infer_semantics(profile)
    intent = _classify_intent(question, profile)  # classify once, reuse across retries

    error = None

    for attempt in range(3):
        logger.info("=" * 60)
        logger.info(f"[ATTEMPT {attempt + 1}/3]")
        try:
            plan = _plan(question, profile, semantics, error)
            plan = _sanitize_plan(plan)

            if not plan:
                plan = _default_plan(df)

            result = _execute(plan, df)
            result = _force_aggregate_if_needed(intent, result)  # LLM-driven safety net
            answer = _respond(question, result, semantics=semantics)

            logger.info("=" * 60)
            logger.info(f"[AGENT SUCCESS] Answer: {answer}")
            return answer

        except Exception as e:
            logger.error(f"[ATTEMPT {attempt + 1}] FAILED: {e}")
            error = str(e)

    logger.error("[AGENT FAILED] All 3 attempts exhausted")
    raise HTTPException(500, "Agent failed after retries")


# =========================
# Backward compatibility
# =========================
def analyze_dataframe(question: str, df: pd.DataFrame) -> str:
    return analyze_dataframe_agent(question, df)