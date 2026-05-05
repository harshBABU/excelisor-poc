# aibackend/smart_router.py
"""
Smart HITL router.

The router uses the LLM to decide which action choices to show before any
Excel operation runs. The selected option's mode is later used by /smart_route
to dispatch the execution handler.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


CONFIDENCE_AUTO_EXECUTE = 0.80
CONFIDENCE_SHOW_RATIONALE = 0.55

MODES = ["ask_ai", "ai_chart", "ai_action_formula", "python_analysis", "audit"]

MODE_DESCRIPTIONS = """
AVAILABLE MODES - choose from these only:

1. ask_ai
   Use for: open-ended insight, trends, explanations, comparisons, forecasts, concise answers in the task pane.

2. ai_chart
   Use for: visualizing data as a chart, graph, plot, dashboard, or month-over-month view.

3. ai_action_formula
   Use for: simple groupby aggregations written into Excel as native formulas.

4. python_analysis
   Use for: top-N, bottom-N, Pareto, cumulative percent, growth, ranking, or multi-step computed analysis.

5. audit
   Use for: data quality, missing values, duplicates, validation, cleaning, and audit tasks.
"""


@dataclass
class ActionOption:
    id: str
    label: str
    mode: str
    description: Optional[str] = None


@dataclass
class RouterResult:
    mode: str
    confidence: float
    rationale: str
    needs_clarification: bool
    clarification_question: Optional[str] = None
    clarification_options: List[str] = field(default_factory=list)
    action_options: List[ActionOption] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)


def _get_feasible_modes(question: str, df: pd.DataFrame) -> List[str]:
    feasible = ["ask_ai"]
    
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    date_cols = [c for c in df.columns if any(k in c.lower() for k in ["date", "month", "year", "time"])]
    has_numeric = len(numeric_cols) >= 1
    has_groupable = len(cat_cols) >= 1 or len(date_cols) >= 1
    enough_rows = len(df) >= 10

    if has_numeric and has_groupable:
        feasible.append("ai_chart")
        feasible.append("ai_action_formula")

    if has_numeric and enough_rows:
        feasible.append("python_analysis")

    feasible.append("audit")
    return feasible


def _build_system_prompt(feasible_modes: List[str]) -> str:
    return f"""You are the Human-in-the-Loop action planner for an AI-powered Excel assistant.
Your job is to parse the user's request and decide what clickable action choices to show BEFORE anything executes.

{MODE_DESCRIPTIONS}

Rules:
- Return ONLY valid JSON, no markdown.
- Always generate 2 to 4 action_options.
- Put the most relevant option first.
- Each option must be a meaningfully different outcome.
- Each option must include id, label, mode, and description.
- id must be stable snake_case.
- label must be short, imperative, and button-friendly.
- YOU MUST ONLY use modes from this allowed list: {feasible_modes}. Never use modes outside this list.
- Parameter Clarification: If the user selects a mode that requires a parameter (e.g., 'N' for Top-N or Bottom-N python_analysis) AND the value is not in the question or PREVIOUS CONTEXT:
  1. Set confidence < 0.55 to trigger parameter clarification.
  2. Set clarification_question to ask for the missing parameter (e.g., "How many items would you like to see?").
  3. Provide action_options for the parameter choices (e.g., labels: "Top 5", "Top 10", "Top 20", "All"), keeping the mode the same.
- Include a concise rationale for your top recommended mode.
- Include a short clarification_question that asks the user to pick an action (if not asking for a parameter).
- Set clarification_options to the action option labels for backward compatibility.
"""


def _build_user_prompt(question: str, df: pd.DataFrame, context: Optional[str]) -> str:
    columns = df.columns.tolist()
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    sample = df.head(3).to_csv(index=False)
    context_section = f"\n\nPREVIOUS CONTEXT:\n{context}" if context else ""

    return f"""USER QUESTION: {question}

DATA PROFILE:
- Shape: {df.shape[0]} rows x {df.shape[1]} columns
- All columns: {columns}
- Numeric columns: {numeric_cols}
- Categorical columns: {categorical_cols}
- Sample (first 3 rows):
{sample}{context_section}

Respond ONLY with this JSON shape:
{{
  "mode": "one_of_the_modes",
  "confidence": 0.0,
  "rationale": "Why the first option is most relevant",
  "clarification_question": "How would you like me to handle this?",
  "clarification_options": ["Option label", "Option label"],
  "action_options": [
    {{
      "id": "stable_snake_case_id",
      "label": "Short clickable label",
      "mode": "one_of_the_modes",
      "description": "One short sentence explaining what will happen."
    }}
  ]
}}"""


def route(question: str, df: pd.DataFrame, context: Optional[str] = None) -> RouterResult:
    feasible_modes = _get_feasible_modes(question, df)

    try:
        from llm_config import get_openai_client, get_model
        client = get_openai_client()
    except Exception as e:
        logger.warning("[SmartRouter] Could not get LLM client, using fallback. Error: %s", e)
        return _fallback_route(question, feasible_modes)

    if client is None:
        return _fallback_route(question, feasible_modes)

    logger.info("[SmartRouter] Planning HITL options for question: '%s...' (feasible: %s)", question[:80], feasible_modes)

    try:
        resp = client.chat.completions.create(
            model=get_model(),
            messages=[
                {"role": "system", "content": _build_system_prompt(feasible_modes)},
                {"role": "user", "content": _build_user_prompt(question, df, context)},
            ],
            response_format={"type": "json_object"},
        )
        raw = (resp.choices[0].message.content or "").strip()
        logger.info("[SmartRouter] LLM raw response: %s", raw[:300])
        data = json.loads(raw)
    except Exception as e:
        logger.error("[SmartRouter] LLM option planning failed: %s", e)
        return _fallback_route(question)

    mode = data.get("mode", "ask_ai")
    if mode not in MODES:
        logger.warning("[SmartRouter] Unknown mode '%s', defaulting to ask_ai", mode)
        mode = "ask_ai"

    confidence = _safe_float(data.get("confidence"), 0.5)
    rationale = str(data.get("rationale") or "")
    action_options = _normalize_action_options(data.get("action_options"), mode, feasible_modes)
    clarification_question = str(data.get("clarification_question") or "How would you like me to handle this?")
    clarification_options = data.get("clarification_options") or [opt.label for opt in action_options]

    return RouterResult(
        mode=mode,
        confidence=confidence,
        rationale=rationale,
        needs_clarification=True,
        clarification_question=clarification_question,
        clarification_options=[str(opt) for opt in clarification_options][:4],
        action_options=action_options,
    )


def _normalize_action_options(raw_options: Any, preferred_mode: str, feasible_modes: List[str]) -> List[ActionOption]:
    options: List[ActionOption] = []
    seen_ids = set()

    if isinstance(raw_options, list):
        for raw in raw_options:
            if not isinstance(raw, dict):
                continue
            mode = raw.get("mode")
            label = str(raw.get("label") or "").strip()
            if mode not in feasible_modes or not label:
                continue

            opt_id = _slugify(str(raw.get("id") or label))
            if opt_id in seen_ids:
                continue
            seen_ids.add(opt_id)
            options.append(ActionOption(
                id=opt_id,
                label=label[:80],
                mode=mode,
                description=str(raw.get("description") or "").strip()[:160] or None,
            ))

    if len(options) < 2:
        return _fallback_options(preferred_mode, "", feasible_modes)

    return options[:4]


def _fallback_route(question: str, feasible_modes: List[str]) -> RouterResult:
    q = question.lower()
    if any(k in q for k in ["chart", "plot", "graph", "visual", "show me", "month"]) and "ai_chart" in feasible_modes:
        mode, conf = "ai_chart", 0.70
    elif any(k in q for k in ["audit", "quality", "missing", "duplicate", "validate"]) and "audit" in feasible_modes:
        mode, conf = "audit", 0.75
    elif any(k in q for k in ["80%", "pareto", "80-20", "top ", "bottom ", "cumulative", "running total", "rank"]) and "python_analysis" in feasible_modes:
        mode, conf = "python_analysis", 0.72
    elif any(k in q for k in ["sum", "total", "count", "average", "by region", "by category", "group"]) and "ai_action_formula" in feasible_modes:
        mode, conf = "ai_action_formula", 0.65
    else:
        mode, conf = "ask_ai", 0.60

    options = _fallback_options(mode, q, feasible_modes)
    return RouterResult(
        mode=mode,
        confidence=conf,
        rationale=f"Fallback action plan led with {mode}.",
        needs_clarification=True,
        clarification_question="How would you like me to handle this?",
        clarification_options=[opt.label for opt in options],
        action_options=options,
    )


def _fallback_options(preferred_mode: str, question: str, feasible_modes: List[str]) -> List[ActionOption]:
    base = {
        "ask_ai": ActionOption("summary_ui", "View summary in add-in", "ask_ai", "Show a concise answer in the task pane."),
        "ai_chart": ActionOption("generate_chart", "Generate chart", "ai_chart", "Create the most suitable chart in Excel."),
        "ai_action_formula": ActionOption("write_formula_summary", "Write formula summary", "ai_action_formula", "Create a worksheet with native Excel formulas."),
        "python_analysis": ActionOption("create_ranked_analysis", "Create ranked analysis", "python_analysis", "Run a computed analysis and write results to Excel."),
        "audit": ActionOption("audit_data_quality", "Audit data quality", "audit", "Check missing values, duplicates, and data quality issues."),
    }

    ordered_modes = [preferred_mode]
    if any(k in question for k in ["chart", "plot", "graph", "visual", "show me", "month"]):
        ordered_modes += ["ai_chart", "ask_ai", "ai_action_formula"]
    elif any(k in question for k in ["top ", "bottom ", "rank", "80%", "pareto"]):
        ordered_modes += ["python_analysis", "ai_chart", "ask_ai"]
    elif any(k in question for k in ["audit", "quality", "missing", "duplicate"]):
        ordered_modes += ["audit", "ask_ai", "ai_action_formula"]
    else:
        ordered_modes += ["ask_ai", "ai_chart", "ai_action_formula"]

    deduped = []
    for mode in ordered_modes:
        if mode in base and mode not in deduped and mode in feasible_modes:
            deduped.append(mode)

    return [base[mode] for mode in deduped[:3]]


def _slugify(value: str) -> str:
    chars = []
    previous_underscore = False
    for ch in value.lower():
        if ch.isalnum():
            chars.append(ch)
            previous_underscore = False
        elif not previous_underscore:
            chars.append("_")
            previous_underscore = True
    return "".join(chars).strip("_")[:48] or "action"


def _safe_float(value: Any, fallback: float) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return fallback
    return max(0.0, min(1.0, numeric))
