# aibackend/main.py

import logging
import os
from typing import Any, Dict, List
import pandas as pd
from contextlib import asynccontextmanager
from types import SimpleNamespace

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from logging_config import configure_logging

configure_logging()
logger = logging.getLogger(__name__)

# Local modules (ensure these files exist in aibackend/)
from models import (
    TablePayload,
    AnalyzeRequest,
    AnalyzeResponse,
    AuditResponse,
    ChartSuggestResponse,
    IntentResponse,
    ActionResponse,
    AiChartRequest,
    SmartRouteRequest,
    SmartRouteResponse,
)
from table_utils import table_to_dataframe, to_python_scalars
from intent import classify_intent, get_python_execution_plan, get_python_execution_plan_from_intent, get_intent_classifier, classify_intent_enhanced
from analysis import analyze_dataframe
from audit import audit_dataframe
from charting import suggest_chart
from worksheet_manager import WorksheetManager
from chart_decider import decide_chart
from data_sampler import IntelligentSampler


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    logger.info("🚀 AI Backend starting up...")
    yield
    # Shutdown  
    logger.info("🛑 AI Backend shutting down gracefully...")

app = FastAPI(lifespan=lifespan)

# CORS: Open for local development; restrict in production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "https://localhost:3000"],  # Allow both HTTP and HTTPS for localhost:3000
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def extract_api_key_and_model(request: Request, call_next):
    from llm_config import api_key_ctx, model_ctx
    auth_header = request.headers.get("Authorization")
    if auth_header and auth_header.startswith("Bearer "):
        api_key_ctx.set(auth_header.split("Bearer ", 1)[1])
        
    model_header = request.headers.get("X-OpenAI-Model")
    if model_header:
        model_ctx.set(model_header)
        
    response = await call_next(request)
    return response

@app.get("/validate-key")
def validate_api_key():
    from llm_config import get_openai_client
    try:
        client = get_openai_client()
        if not client:
            return {"valid": False, "message": "No API key provided."}
        # Single cheap GET to authenticate — much faster than models.list()
        client.models.retrieve("gpt-3.5-turbo")
        return {"valid": True, "message": "API key is valid."}
    except Exception as e:
        err = str(e)
        if "401" in err or "Incorrect API key" in err or "invalid_api_key" in err:
            return {"valid": False, "message": "Invalid API key. Please check and try again."}
        logger.warning(f"Key validation failed: {err}")
        return {"valid": False, "message": "Validation failed. Is the backend connected to the internet?"}

wm = WorksheetManager()
ds = IntelligentSampler()


def execute_python_only_analysis(df: pd.DataFrame, execution_plan: Dict[str, Any]) -> Dict[str, Any]:
    """
    Execute complete Python-only analysis for top N/bottom N queries.
    Returns final results ready for direct worksheet population (no Excel formulas needed).
    """
    logger.info("[PYTHON-ONLY] Starting complete Python analysis")
    logger.debug("[PYTHON-ONLY] execution_plan=%s", execution_plan)
    
    ranking_criteria = execution_plan.get('ranking_criteria')
    filter_operation = execution_plan.get('filter_operation')
    filter_value = execution_plan.get('filter_value')
    final_aggregation = execution_plan.get('final_aggregation')
    aggregation_column = execution_plan.get('aggregation_column')
    groupby_column = execution_plan.get('groupby_column')
    
    logger.info("[PYTHON-ONLY] Parameters: %s %s %s", filter_operation, filter_value, ranking_criteria)
    logger.info("[PYTHON-ONLY] Aggregation: %s of %s by %s", final_aggregation, aggregation_column, groupby_column)
    
    # For top N/bottom N queries with grouping, we need to aggregate first, then filter
    if final_aggregation and aggregation_column and groupby_column and filter_operation in ["top_n", "bottom_n"]:
        logger.info("[PYTHON-ONLY] Step 1: Aggregating %s of %s by %s", final_aggregation, aggregation_column, groupby_column)
        
        agg_func = final_aggregation.lower()
        
        # Apply aggregation first
        if agg_func in ["average", "avg", "mean"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].mean()
        elif agg_func in ["sum", "total"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].sum()
        elif agg_func in ["max", "maximum"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].max()
        elif agg_func in ["min", "minimum"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].min()
        elif agg_func in ["count", "frequency"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].count()
        elif agg_func in ["median", "middle"]:
            aggregated_series = df.groupby(groupby_column)[aggregation_column].median()
        else:
            logger.warning("[PYTHON-ONLY] Unknown aggregation: %s, using sum as fallback", final_aggregation)
            aggregated_series = df.groupby(groupby_column)[aggregation_column].sum()
        
        # Convert to DataFrame for easier manipulation
        aggregated_df = aggregated_series.reset_index()
        aggregated_df.columns = [groupby_column, f'{final_aggregation}_{aggregation_column}']
        
        logger.info("[PYTHON-ONLY] After aggregation: %d groups", len(aggregated_df))
        
        # Step 2: Apply top N/bottom N filtering on aggregated results
        ranking_column = f'{final_aggregation}_{aggregation_column}'
        if filter_operation == "top_n":
            logger.info("[PYTHON-ONLY] Step 2: Filtering to top %s groups by %s", filter_value, ranking_column)
            final_df = aggregated_df.nlargest(filter_value, ranking_column)
        elif filter_operation == "bottom_n":
            logger.info("[PYTHON-ONLY] Step 2: Filtering to bottom %s groups by %s", filter_value, ranking_column)
            final_df = aggregated_df.nsmallest(filter_value, ranking_column)
        else:
            final_df = aggregated_df
        
        logger.info("[PYTHON-ONLY] Final results: %d groups", len(final_df))
        
        # Convert to results format
        results_data = final_df.to_dict('records')
        
        return {
            "type": "aggregated_results",
            "data": results_data,
            "summary": f"Top {filter_value} {groupby_column}s by {final_aggregation} of {aggregation_column}",
            "metadata": {
                "original_rows": len(df),
                "aggregated_groups": len(aggregated_df),
                "final_groups": len(final_df),
                "operation": f"{filter_operation} {filter_value} by {final_aggregation} of {aggregation_column}"
            }
        }
    
    # Step 1: Apply filtering/ranking (for non-aggregated queries)
    if filter_operation == "top_n" and ranking_criteria and filter_value:
        logger.info("[PYTHON-ONLY] Filtering to top %s by %s", filter_value, ranking_criteria)
        filtered_df = df.nlargest(filter_value, ranking_criteria)
    elif filter_operation == "bottom_n" and ranking_criteria and filter_value:
        logger.info("[PYTHON-ONLY] Filtering to bottom %s by %s", filter_value, ranking_criteria)
        filtered_df = df.nsmallest(filter_value, ranking_criteria)
    elif filter_operation == "threshold" and ranking_criteria and filter_value:
        logger.info("[PYTHON-ONLY] Filtering where %s > %s", ranking_criteria, filter_value)
        filtered_df = df[df[ranking_criteria] > filter_value]
    else:
        logger.info("[PYTHON-ONLY] No filtering applied")
        filtered_df = df
    
    logger.info("[PYTHON-ONLY] After filtering: %d rows", filtered_df.shape[0])
    
    # Step 2: Apply aggregation if specified (for non-top N queries)
    if final_aggregation and aggregation_column and groupby_column:
        logger.info("[PYTHON-ONLY] Applying %s on %s grouped by %s", final_aggregation, aggregation_column, groupby_column)
        
        agg_func = final_aggregation.lower()
        
        # Basic Aggregations
        if agg_func in ["average", "avg", "mean"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].mean()
        elif agg_func in ["sum", "total"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].sum()
        elif agg_func in ["max", "maximum"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].max()
        elif agg_func in ["min", "minimum"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].min()
        elif agg_func in ["count", "frequency"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].count()
            
        # Statistical Aggregations
        elif agg_func in ["median", "middle"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].median()
        elif agg_func in ["std", "stdev", "standard_deviation"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].std()
        elif agg_func in ["var", "variance"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].var()
        elif agg_func in ["sem", "standard_error"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].sem()
        elif agg_func in ["mad", "mean_absolute_deviation"]:
            # MAD calculation: mean of absolute deviations from the mean
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: (x - x.mean()).abs().mean())
            
        # Percentiles and Quantiles
        elif agg_func in ["q1", "first_quartile", "25th_percentile"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.25)
        elif agg_func in ["q3", "third_quartile", "75th_percentile"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.75)
        elif agg_func in ["iqr", "interquartile_range"]:
            q1 = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.25)
            q3 = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.75)
            result_series = q3 - q1
        elif agg_func in ["90th_percentile", "p90"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.9)
        elif agg_func in ["95th_percentile", "p95"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.95)
        elif agg_func in ["99th_percentile", "p99"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].quantile(0.99)
            
        # Range and Spread
        elif agg_func in ["range", "spread"]:
            max_vals = filtered_df.groupby(groupby_column)[aggregation_column].max()
            min_vals = filtered_df.groupby(groupby_column)[aggregation_column].min()
            result_series = max_vals - min_vals
        elif agg_func in ["unique_count", "nunique", "distinct_count"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].nunique()
            
        # Growth and Change (requires sorted data)
        elif agg_func in ["first", "earliest"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].first()
        elif agg_func in ["last", "latest"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].last()
        elif agg_func in ["cumulative_sum", "cumsum"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].cumsum().groupby(filtered_df[groupby_column]).last()
            
        # Business Metrics
        elif agg_func in ["coefficient_of_variation", "cv"]:
            means = filtered_df.groupby(groupby_column)[aggregation_column].mean()
            stds = filtered_df.groupby(groupby_column)[aggregation_column].std()
            result_series = stds / means
        elif agg_func in ["skewness", "skew"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].skew()
        elif agg_func in ["kurtosis", "kurt"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: x.kurtosis())
            
        # Custom Aggregations
        elif agg_func in ["positive_count"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: (x > 0).sum())
        elif agg_func in ["negative_count"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: (x < 0).sum())
        elif agg_func in ["zero_count"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: (x == 0).sum())
        elif agg_func in ["non_zero_count"]:
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].apply(lambda x: (x != 0).sum())
            
        else:
            logger.warning("[PYTHON EXEC] Unknown aggregation: %s, using mean as fallback", final_aggregation)
            result_series = filtered_df.groupby(groupby_column)[aggregation_column].mean()
        
        # Convert to DataFrame
        result_df = result_series.reset_index()
        result_df.columns = [groupby_column, f'{final_aggregation}_{aggregation_column}']
        
        logger.info("[PYTHON EXEC] After aggregation: %d rows, %d columns", result_df.shape[0], result_df.shape[1])
        return result_df
    else:
        logger.info("[PYTHON EXEC] No aggregation applied, returning filtered data")
        return filtered_df



# -------------- Smart Intent Router --------------

MAX_CLARIFICATION_ROUNDS = 2

@app.post("/smart_route", response_model=SmartRouteResponse)
def smart_route(req: SmartRouteRequest) -> SmartRouteResponse:
    """
    Single entry point that:
    1. Classifies the question into an execution mode via SmartRouter.
    2. If confidence is low AND clarification rounds remaining → return HITL options.
    3. Otherwise → dispatch to the appropriate sub-handler and embed results.
    """
    from smart_router import route as sr_route, CONFIDENCE_SHOW_RATIONALE

    df = table_to_dataframe(req.table)
    logger.info(
        "[/smart_route] question='%s' selected_action=%s selected_mode=%s",
        req.question[:80],
        req.selected_action_id,
        req.selected_mode,
    )

    if req.selected_mode:
        result = SimpleNamespace(
            mode=req.selected_mode,
            confidence=1.0,
            rationale=f"User selected action: {req.selected_action_id or req.selected_mode}",
            needs_clarification=False,
            clarification_question=None,
            clarification_options=[],
            action_options=[],
        )
    else:
        result = sr_route(req.question, df, context=req.context)

    # Force execution after max clarification rounds to avoid infinite loops
    force_execute = req.clarification_round >= MAX_CLARIFICATION_ROUNDS
    if force_execute and result.needs_clarification:
        logger.info("[/smart_route] Max clarification rounds reached — forcing execution")
        result.needs_clarification = False
        result.clarification_question = None
        result.clarification_options = []

    # If we still need clarification, return early with HITL payload
    if result.needs_clarification and not req.selected_mode:
        return SmartRouteResponse(
            mode=result.mode,
            confidence=result.confidence,
            rationale=result.rationale,
            needs_clarification=True,
            clarification_question=result.clarification_question,
            clarification_options=result.clarification_options,
            action_options=[
                {
                    "id": opt.id,
                    "label": opt.label,
                    "mode": opt.mode,
                    "description": opt.description,
                }
                for opt in result.action_options
            ],
        )

    if req.selected_mode:
        if req.selected_mode not in ("ask_ai", "ai_chart", "ai_action_formula", "python_analysis", "audit"):
            raise HTTPException(status_code=400, detail=f"Unsupported selected_mode: {req.selected_mode}")
        result.mode = req.selected_mode
        result.needs_clarification = False
        result.clarification_question = None
        result.clarification_options = []
        result.action_options = []

    # ── Dispatch to sub-handler ──────────────────────────────────────────────
    show_rationale = CONFIDENCE_SHOW_RATIONALE <= result.confidence < 0.80
    rationale_prefix = f"💡 I interpreted this as: {result.rationale}\n\n" if show_rationale else ""

    mode = result.mode

    # ── ask_ai ───────────────────────────────────────────────────────────────
    if mode == "ask_ai":
        try:
            answer = analyze_dataframe(req.question, df)
            return SmartRouteResponse(
                mode=mode, confidence=result.confidence, rationale=result.rationale,
                needs_clarification=False,
                analyze_result=AnalyzeResponse(answer=rationale_prefix + answer),
            )
        except Exception as e:
            return SmartRouteResponse(
                mode=mode, confidence=result.confidence, rationale=result.rationale,
                needs_clarification=False,
                analyze_result=AnalyzeResponse(answer=f"❌ Analysis failed: {e}"),
            )

    # ── ai_chart ─────────────────────────────────────────────────────────────
    if mode == "ai_chart":
        chart_type, chart_title, cat_col, val_cols, _, inferences = decide_chart(df, req.question, True)
        processed_value_columns, processed_aggregations = [], []
        if isinstance(val_cols, dict):
            for _, cfg in val_cols.items():
                processed_value_columns.append(cfg["column"])
                processed_aggregations.append(cfg["aggregation"])
        elif isinstance(val_cols, list):
            processed_value_columns = val_cols
            processed_aggregations = ["SUM"] * len(val_cols)
        category_columns = cat_col if isinstance(cat_col, (list, type(None))) else [cat_col]
        return SmartRouteResponse(
            mode=mode, confidence=result.confidence, rationale=rationale_prefix + result.rationale,
            needs_clarification=False,
            chart_result=ChartSuggestResponse(
                chartType=chart_type, chartTitle=chart_title,
                categoryColumn=category_columns,
                valueColumns=processed_value_columns,
                aggregations=processed_aggregations,
                inferences=inferences,
            ),
        )

    # ── audit ────────────────────────────────────────────────────────────────
    if mode == "audit":
        audit_res = audit_dataframe(df)
        return SmartRouteResponse(
            mode=mode, confidence=result.confidence, rationale=result.rationale,
            needs_clarification=False,
            action_result=ActionResponse(**audit_res),
        )

    # ── ai_action_formula & python_analysis ──────────────────────────────────
    # Both go through the existing /action orchestration logic
    if mode in ("ai_action_formula", "python_analysis"):
        # Re-use existing action_route logic by constructing a fake request
        from models import AnalyzeRequest as _AR
        fake_req = _AR(table=req.table, address=req.address, question=req.question)
        action_resp = action_route(fake_req)
        # Prefix rationale if borderline confidence
        action_resp.message = rationale_prefix + action_resp.message
        return SmartRouteResponse(
            mode=mode, confidence=result.confidence, rationale=result.rationale,
            needs_clarification=False,
            action_result=action_resp,
        )

    # Catch-all fallback
    return SmartRouteResponse(
        mode="ask_ai", confidence=0.5, rationale="Unknown mode — falling back to text analysis",
        needs_clarification=False,
        analyze_result=AnalyzeResponse(answer="Could not determine the right action. Please rephrase."),
    )


# -------------- Health --------------

@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


# -------------- Intent --------------

@app.post("/classify_intent", response_model=IntentResponse)
def classify_intent_route(req: AnalyzeRequest) -> IntentResponse:
    df = table_to_dataframe(req.table)
    result = classify_intent(req.question, df)
    return IntentResponse(**to_python_scalars(result))


# -------------- Analyze --------------

@app.post("/analyze", response_model=AnalyzeResponse)
def analyze_route(req: AnalyzeRequest) -> AnalyzeResponse:
    df = table_to_dataframe(req.table)
    try:
        answer = analyze_dataframe(req.question, df)
    except HTTPException as e:
        # Bubble up expected HTTP error codes (e.g., missing question)
        raise e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Analyze failed: {e}")
    return AnalyzeResponse(answer=answer)


# -------------- Audit --------------

@app.post("/ai_audit", response_model=ActionResponse)
def ai_audit_route(req: TablePayload) -> ActionResponse:
    df = table_to_dataframe(req.table)
    result = audit_dataframe(df)
    return ActionResponse(**result)


# -------------- Chart Suggestion --------------

@app.post("/ai_chart", response_model=ChartSuggestResponse)
def ai_chart_route(req: AiChartRequest) -> ChartSuggestResponse:
    df = table_to_dataframe(req.table)
    chart_type, chart_title, cat_col, val_cols, _, inferences = decide_chart(df, req.question, bool(req.use_llm))
    
    # Process the new valueColumns format
    processed_value_columns = []
    processed_aggregations = []
    
    if isinstance(val_cols, dict):
        # New format: {"series_name": {"column": "col", "aggregation": "AGG"}}
        for series_name, series_config in val_cols.items():
            column = series_config["column"]
            aggregation = series_config["aggregation"]
            processed_value_columns.append(column)
            processed_aggregations.append(aggregation)
    elif isinstance(val_cols, list):
        # Legacy format: handle gracefully
        processed_value_columns = val_cols
        processed_aggregations = ["SUM"] * len(val_cols)
    else:
        processed_value_columns = []
        processed_aggregations = []
    
    # Handle category column(s)
    category_columns = cat_col if isinstance(cat_col, (list, type(None))) else [cat_col]
    
    logger.debug("[Backend] Category columns: %s", category_columns)
    logger.debug("[Backend] Returning aggregations: %s", processed_aggregations)
    
    response = ChartSuggestResponse(
        chartType=chart_type,
        chartTitle=chart_title,
        categoryColumn=category_columns,  # Now accepts list or single value
        valueColumns=processed_value_columns,
        aggregations=processed_aggregations,
        inferences=inferences
    )
    logger.info("[Backend] Chart response: %s", response)
    return response



# -------------- Action Orchestration --------------

@app.post("/action", response_model=ActionResponse)
def action_route(req: AnalyzeRequest) -> ActionResponse:
    """
    Orchestrate the requested action based on explicit "formula" hint or classified intent.
    
    Behavior:
      - If the question mentions "formula/formulas": build a formula sheet from the selection.
      - If intent suggests ranking/top analysis: create formula-driven analysis sheet
      - If intent == "chart": tell the frontend to call /ai_chart and insert the chart
      - If intent == "audit": build a data quality worksheet payload
      - If intent == "analyze": run LLM-backed analysis and return a worksheet
    """
    df = table_to_dataframe(req.table)
    logger.debug("Table in mainpy: length=%d", len(req.table))
    logger.debug("DataFrame in mainpy: length=%d", len(df))
    
    q_lower = (req.question or "").lower()
    
    # PRIORITY 1: Enhanced intent-based routing (handles top N/bottom N with Python-only approach)
    # Recent fix: Use enhanced classifier directly to get both intent and analysis plan
    classifier = get_intent_classifier()
    intent_result = classifier.classify_intent(req.question, df)
    intent = {
        "action_type": intent_result.action_type.value,
        "confidence": intent_result.confidence,
        "entities": intent_result.entities
    }
    
    # Check for top N/bottom N queries first and route them to Python-only handler
    # Recent fix: Reuse intent_result to avoid redundant classification
    python_plan = get_python_execution_plan_from_intent(intent_result)
    logger.debug("[DEBUG] Python plan check: requires_preprocessing=%s", python_plan['requires_python_preprocessing'])
    if python_plan["requires_python_preprocessing"]:
        execution_plan = python_plan.get('execution_plan', {})
        if execution_plan is None:
            execution_plan = {}
        filter_operation = execution_plan.get('filter_operation')
        logger.debug("[DEBUG] Filter operation detected: %s", filter_operation)
        
        # Route top N/bottom N queries to Python-only handler immediately
        if filter_operation in ["top_n", "bottom_n"]:
            logger.info("[PYTHON-ONLY] Early routing %s query to Python-only handler - RETURNING IMMEDIATELY", filter_operation)
            
            # Execute complete Python analysis
            python_results = execute_python_only_analysis(df, execution_plan)
            
            # Create Python-only results sheet (no Excel formulas)
            sheet = wm.make_python_results_sheet(python_results, hint=req.question or "")
            
            return ActionResponse(
                success=True,
                message=f"Python-only analysis completed: {python_results.get('summary', 'Results ready')}",
                worksheets=[sheet],
                formulas=[]  # No formulas - all values pre-calculated in Python
            )
        else:
            logger.debug("[DEBUG] Not a top_n/bottom_n query, filter_operation=%s, continuing to other handlers", filter_operation)
    else:
        logger.debug("[DEBUG] No Python preprocessing required, continuing to formula/summary handlers")
    
    # PRIORITY 2: Enhanced formula request detection (only for non-top N queries)
    formula_keywords = ["formula", "formulas", "rank", "highest", "lowest", "sort", "order by"]  # Removed "top" to avoid conflicts
    is_formula_request = any(keyword in q_lower for keyword in formula_keywords)
    
    # Recent fix: Detect summary/aggregation requests (but with lower priority than Python-only analysis)
    summary_keywords = ["unique", "sum", "total", "count", "group", "aggregate", "summary", "by sku", "by category", "distinct"]
    is_summary_request = any(keyword in q_lower for keyword in summary_keywords)
    
    # Only process as summary request if it's NOT a top N/bottom N query (those should use Python-only)
    if is_summary_request and not (python_plan["requires_python_preprocessing"] and python_plan.get('execution_plan', {}).get('filter_operation') in ["top_n", "bottom_n"]):
        # Recent fix: Use FULL dataset for Excel formula generation to avoid sampling mismatches
        logger.debug("[DEBUG] Summary request - Using full dataset (%d rows) for accurate Excel formulas", len(df))
        data_context = ds.sample_large_dataset(df.values.tolist())  # Force full data by converting to list first
        # Override with full data context to ensure Excel formulas reference complete dataset
        data_context["pattern_sample"] = [df.columns.tolist()] + df.values.tolist()
        data_context["full_data_available"] = False  # Prevent further sampling
        data_context["sampling_strategy"] = "full_data_for_formulas"
        
        # For aggregation queries, extract groupby pattern from LLM analysis plan if available
        execution_plan = python_plan.get('execution_plan')
        
        # Recent fix: For non-Python queries, use execution plan from already obtained intent_result
        if not execution_plan:
            logger.debug("[DEBUG] No execution plan from python_plan, checking intent_result analysis plan")
            if intent_result.analysis_plan:
                execution_plan = intent_result.analysis_plan
                logger.debug("[DEBUG] Got execution plan from intent_result analysis: %s", execution_plan)
        
        logger.debug("[DEBUG] Checking LLM execution plan: %s", execution_plan)
        # Recent fix: Support both old and new LLM analysis plan formats
        if execution_plan:
            # Check for new multi-aggregation format first
            aggregation_details = execution_plan.get('aggregation_details', [])
            groupby_columns = execution_plan.get('groupby_columns', [])
            
            if aggregation_details and groupby_columns:
                logger.debug("[DEBUG] Using LLM multi-aggregation plan with %d aggregations and %d groupby columns", len(aggregation_details), len(groupby_columns))
                # Recent fix: Pass full multi-aggregation pattern instead of just first aggregation
                data_context["llm_multi_aggregation_pattern"] = {
                    "aggregation_details": aggregation_details,
                    "groupby_columns": groupby_columns,
                    "is_multi_dimensional": True,
                    "multi_aggregation": True
                }
                logger.debug("[DEBUG] LLM-derived multi-aggregation pattern: %d aggregations on %s", len(aggregation_details), groupby_columns)
                
                # Also set single pattern for backward compatibility with existing logic
                primary_agg = aggregation_details[0]
                data_context["llm_groupby_pattern"] = {
                    "metric_column": primary_agg.get('column'),
                    "groupby_column": groupby_columns[0], 
                    "aggregation": primary_agg.get('function', 'SUM').upper()
                }
                
            # Fallback to old single-aggregation format
            elif (execution_plan.get('final_aggregation') and 
                  execution_plan.get('aggregation_column') is not None and 
                  execution_plan.get('groupby_column') is not None):
                logger.debug("[DEBUG] Using LLM single-aggregation plan (legacy format)")
                data_context["llm_groupby_pattern"] = {
                    "metric_column": execution_plan.get('aggregation_column'),
                    "groupby_column": execution_plan.get('groupby_column'), 
                    "aggregation": execution_plan.get('final_aggregation', 'SUM').upper()
                }
                logger.debug("[DEBUG] LLM-derived groupby pattern: %s", data_context['llm_groupby_pattern'])
            else:
                logger.debug("[DEBUG] LLM execution plan validation failed - insufficient aggregation/groupby information")

        logger.debug("[DEBUG] Summary request - Using %d data rows for formulas", len(data_context['pattern_sample'])-1)
        sheet = wm.make_calculated_values_sheet(data_context, hint=req.question or "")
        # Recent fix: Extract formulas from worksheet instead of hardcoding empty array
        sheet_formulas = sheet.get("formulas", [])
        logger.debug("[DEBUG] Summary request - Created sheet with %d formulas", len(sheet_formulas))
        #print(f"[DEBUG] Summary request - Sheet formulas: {sheet_formulas}")
        return ActionResponse(
            success=True,
            message="Intelligent summary worksheet created with calculated values using full dataset.",
            worksheets=[sheet],
            formulas=sheet_formulas
        )
    
    if "formula" in q_lower or "formulas" in q_lower or is_formula_request:
        # Recent fix: Use FULL dataset for Excel formula generation to avoid sampling mismatches
        logger.debug("[DEBUG] Formula request - Using full dataset (%d rows) for accurate Excel formulas", len(df))
        data_context = ds.sample_large_dataset(df.values.tolist())  # Force full data by converting to list first
        # Override with full data context to ensure Excel formulas reference complete dataset
        data_context["pattern_sample"] = [df.columns.tolist()] + df.values.tolist()
        data_context["full_data_available"] = False  # Prevent further sampling
        data_context["sampling_strategy"] = "full_data_for_formulas"
        sheet = wm.make_calculated_values_sheet(data_context, hint=req.question or "")
        # Recent fix: Extract formulas from worksheet instead of hardcoding empty array
        sheet_formulas = sheet.get("formulas", [])
        logger.debug("[DEBUG] Formula request - Created sheet with %d formulas", len(sheet_formulas))
        # logger.debug("[DEBUG] Formula request - Sheet formulas: %s", sheet_formulas)
        return ActionResponse(
            success=True,
            message="Formula worksheet created with intelligent analysis using full dataset.",
            worksheets=[sheet],
            formulas=sheet_formulas
        )
    
    # Note: Intent classification and Python preprocessing already handled above for top N/bottom N
    # This section handles other action types that don't require Python-only approach
    
    if intent["action_type"] == "chart":
        return ActionResponse(
            success=True,
            message="Chart recommended. Use AI Chart action to insert.",
            worksheets=[],
            formulas=[]
        )
    
    # Recent fix: Add missing audit action handling
    if intent["action_type"] == "audit":
        audit_res = audit_dataframe(df)
        return ActionResponse(**audit_res)
    
    # Recent fix: Add missing analyze action handling  
    if intent["action_type"] == "analyze":
        try:
            answer = analyze_dataframe(req.question, df)
            sheet = wm.make_analysis_sheet(answer)
            # Recent fix: Extract formulas from worksheet instead of hardcoding empty array
            sheet_formulas = sheet.get("formulas", [])
            return ActionResponse(
                success=True,
                message="Analysis completed and worksheet created.",
                worksheets=[sheet],
                formulas=sheet_formulas
            )
        except Exception as e:
            return ActionResponse(
                success=False,
                message=f"Analysis failed: {str(e)}",
                worksheets=[],
                formulas=[]
            )

    # Recent fix: Add missing summary action handling
    if intent["action_type"] == "summary":
        # Use calculated values sheet for summary requests with LLM analysis plan
        logger.debug("[DEBUG] Summary request detected - Using LLM analysis plan for %s", req.question)
        # Recent fix: Preserve original column names to avoid incorrect mapping
        data_context = {
            "pattern_sample": [df.columns.tolist()] + df.values.tolist(),
            "full_data_available": False,
            "sampling_strategy": "full_data_for_formulas"
        }
        
        # Pass the LLM analysis plan to guide formula generation - use original column names
        execution_plan = intent_result.analysis_plan if hasattr(intent_result, 'analysis_plan') else None
        if execution_plan:
            data_context["llm_groupby_pattern"] = {
                "metric_column": execution_plan.get('aggregation_column'),  # Should be 'PR_QTY'
                "groupby_column": execution_plan.get('groupby_column'),     # Should be 'CATEGORY'
                "aggregation": execution_plan.get('final_aggregation', 'SUM').upper()
            }
            logger.debug("[DEBUG] Summary - Using LLM analysis plan with original column names: %s", data_context['llm_groupby_pattern'])
        
        sheet = wm.make_calculated_values_sheet(data_context, hint=req.question or "")
        sheet_formulas = sheet.get("formulas", [])
        logger.debug("[DEBUG] Summary request - Created sheet with %d formulas", len(sheet_formulas))
        
        return ActionResponse(
            success=True,
            message="Summary worksheet created with statistical calculations.",
            worksheets=[sheet],
            formulas=sheet_formulas
        )
    
    # Handle remaining formula action types (non-top N/bottom N)
    if intent["action_type"] == "formula":
        try:
            # Standard formula approach for simple queries that don't need Python-only
            logger.info("[FORMULA] Using standard formula approach for non-top N query with full dataset (%d rows)", len(df))
            # Recent fix: Preserve original column names to avoid incorrect mapping
            data_context = {
                "pattern_sample": [df.columns.tolist()] + df.values.tolist(),
                "full_data_available": False,
                "sampling_strategy": "full_data_for_formulas"
            }
            
            # Pass the LLM analysis plan to guide formula generation - use original column names
            execution_plan = intent_result.analysis_plan if hasattr(intent_result, 'analysis_plan') else None
            if execution_plan:
                data_context["llm_groupby_pattern"] = {
                    "metric_column": execution_plan.get('aggregation_column'),  # Should be 'PR_QTY'
                    "groupby_column": execution_plan.get('groupby_column'),     # Should be 'CATEGORY'
                    "aggregation": execution_plan.get('final_aggregation', 'SUM').upper()
                }
                logger.debug("[DEBUG] Formula - Using LLM analysis plan with original column names: %s", data_context['llm_groupby_pattern'])
            
            sheet = wm.make_calculated_values_sheet(data_context, hint=req.question or "")
            sheet_formulas = sheet.get("formulas", [])
            
            return ActionResponse(
                success=True,
                message="Formula worksheet created with calculated values.",
                worksheets=[sheet],
                formulas=sheet_formulas
            )
                
        except Exception as e:
            logger.error("[ERROR] Formula action failed: %s", str(e))
            import traceback
            traceback.print_exc()
            return ActionResponse(
                success=False,
                message=f"Formula creation failed: {str(e)}",
                worksheets=[],
                formulas=[]
            )

    # Fallback: default action for unrecognized intents
    return ActionResponse(
        success=False,
        message="Could not determine action from the question. Try asking for 'formulas', 'chart', 'audit', or analysis.",
        worksheets=[],
        formulas=[]
    )
