# aibackend/smart_router.py
"""
Smart Intent Router — single LLM call that classifies every user query
into one of 5 execution modes and returns a confidence score.

Modes
-----
ask_ai              Open-ended insight / trend / general question → /analyze
ai_chart            User wants a visualisation                    → /ai_chart
ai_action_formula   Simple groupby aggregation, summary table     → /action (formula path)
python_analysis     Complex: top-N, pareto, cumulative%, growth   → /action (python path)
audit               Data quality, missing values, duplicates      → /ai_audit
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


CONFIDENCE_AUTO_EXECUTE   = 0.80   # execute silently
CONFIDENCE_SHOW_RATIONALE = 0.55   # execute but show "I interpreted this as…"
# Below 0.55 → surface HITL clarification options


MODES = ["ask_ai", "ai_chart", "ai_action_formula", "python_analysis", "audit"]

MODE_DESCRIPTIONS = """
AVAILABLE MODES — choose exactly ONE:

1. ask_ai
   Use for: open-ended questions, general insight, trend explanations, "why did X happen", comparisons, forecasts.
   Examples: "Why did sales drop in Q3?", "What are the key trends?", "Summarise the data for me"

2. ai_chart
   Use for: any request to visualise data — chart, graph, plot, dashboard, visual.
   Examples: "Show month on month sales", "Plot revenue by region as a bar chart", "Give me a pie chart of top products"

3. ai_action_formula
   Use for: simple groupby aggregations that can be expressed as SUMIFS/AVERAGEIFS/COUNTIFS.
   The result is written into Excel as native formulas.
   Examples: "Sum sales by region", "Average price per category", "Count orders by status"

4. python_analysis
   Use for: ANYTHING requiring multi-step computation beyond a single aggregation:
   - Pareto / 80-20 / cumulative percentage
   - Top-N or Bottom-N items WITH a threshold or filter
   - Percentile-based filtering ("items generating 80% of revenue")
   - Running totals, growth rates, period-over-period change
   - Ranking + secondary metric calculation
   Examples: "Which item types generate 80% of revenue?",
             "Top 5 products by sales with average margin",
             "Show cumulative revenue contribution"

5. audit
   Use for: data quality, validation, cleaning tasks.
   Examples: "Check for missing values", "Find duplicates", "Validate data quality", "Audit the dataset"
"""


@dataclass
class RouterResult:
    mode: str                          # one of MODES
    confidence: float                  # 0.0 – 1.0
    rationale: str                     # human-readable reason
    needs_clarification: bool          # True if confidence < threshold
    clarification_question: Optional[str] = None   # what to ask the user
    clarification_options: List[str] = field(default_factory=list)  # clickable option labels
    metadata: Dict[str, Any] = field(default_factory=dict)          # hints for downstream handlers


def _build_system_prompt() -> str:
    return f"""You are the Smart Intent Router for an AI-powered Excel assistant.
Your ONLY job is to classify the user's question into ONE execution mode and return a confidence score.

{MODE_DESCRIPTIONS}

RULES:
- Return ONLY valid JSON — no markdown, no explanation outside the JSON.
- "mode" MUST be one of: {MODES}
- "confidence" MUST be a float between 0.0 and 1.0
- "clarification_question" — if you are unsure (confidence < 0.55), write a SHORT question to ask the user
- "clarification_options" — 2 to 4 short option labels matching the ambiguous modes (e.g. ["Show a chart", "Give me a table", "Text analysis"])
- If confidence >= 0.55, set clarification_question to null and clarification_options to []
"""


def _build_user_prompt(question: str, df: pd.DataFrame, context: Optional[str]) -> str:
    columns = df.columns.tolist()
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    sample = df.head(3).to_csv(index=False)

    context_section = f"\n\nPREVIOUS CONTEXT:\n{context}" if context else ""

    return f"""USER QUESTION: {question}

DATA PROFILE:
- Shape: {df.shape[0]} rows × {df.shape[1]} columns
- All columns: {columns}
- Numeric columns: {numeric_cols}
- Categorical columns: {categorical_cols}
- Sample (first 3 rows):
{sample}{context_section}

Respond ONLY with this JSON structure:
{{
  "mode": "one_of_the_five_modes",
  "confidence": 0.0_to_1.0,
  "rationale": "1-2 sentence explanation of why you chose this mode",
  "clarification_question": "question to ask user OR null",
  "clarification_options": ["Option A", "Option B"] // or []
}}"""


def route(question: str, df: pd.DataFrame, context: Optional[str] = None) -> RouterResult:
    """
    Classify the user's question into an execution mode.

    Parameters
    ----------
    question : str
        The raw user question.
    df : pd.DataFrame
        The data currently selected in Excel.
    context : str, optional
        Prior conversation context or user clarification answer.

    Returns
    -------
    RouterResult
    """
    try:
        from llm_config import get_openai_client, get_model
        client = get_openai_client()
    except Exception as e:
        logger.warning(f"[SmartRouter] Could not get LLM client, using keyword fallback. Error: {e}")
        return _fallback_route(question)

    if client is None:
        return _fallback_route(question)

    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(question, df, context)

    logger.info(f"[SmartRouter] Routing question: '{question[:80]}...'")

    try:
        resp = client.chat.completions.create(
            model=get_model(),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
        )
        raw = (resp.choices[0].message.content or "").strip()
        logger.info(f"[SmartRouter] LLM raw response: {raw[:300]}")
        data = json.loads(raw)
    except Exception as e:
        logger.error(f"[SmartRouter] LLM call failed: {e}")
        return _fallback_route(question)

    mode = data.get("mode", "ask_ai")
    if mode not in MODES:
        logger.warning(f"[SmartRouter] Unknown mode '{mode}' — defaulting to ask_ai")
        mode = "ask_ai"

    confidence = float(data.get("confidence", 0.5))
    rationale = data.get("rationale", "")
    clarification_question = data.get("clarification_question")
    clarification_options = data.get("clarification_options", [])

    needs_clarification = confidence < CONFIDENCE_SHOW_RATIONALE

    result = RouterResult(
        mode=mode,
        confidence=confidence,
        rationale=rationale,
        needs_clarification=needs_clarification,
        clarification_question=clarification_question if needs_clarification else None,
        clarification_options=clarification_options if needs_clarification else [],
    )
    logger.info(f"[SmartRouter] → mode={mode}, confidence={confidence:.2f}, hitl={needs_clarification}")
    return result


def _fallback_route(question: str) -> RouterResult:
    """Keyword-based fallback when LLM is unavailable."""
    q = question.lower()
    if any(k in q for k in ["chart", "plot", "graph", "visual", "show me"]):
        mode, conf = "ai_chart", 0.70
    elif any(k in q for k in ["audit", "quality", "missing", "duplicate", "validate"]):
        mode, conf = "audit", 0.75
    elif any(k in q for k in ["80%", "pareto", "80-20", "top ", "bottom ", "cumulative", "running total"]):
        mode, conf = "python_analysis", 0.72
    elif any(k in q for k in ["sum", "total", "count", "average", "by region", "by category", "group"]):
        mode, conf = "ai_action_formula", 0.65
    else:
        mode, conf = "ask_ai", 0.60

    return RouterResult(
        mode=mode,
        confidence=conf,
        rationale=f"Auto-detected mode: {mode} (keyword match)",
        needs_clarification=False,
    )
