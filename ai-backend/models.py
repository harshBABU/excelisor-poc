# aibackend/models.py

import json
import logging
from typing import Any, Dict, List, Optional, Tuple, Union

from pydantic import BaseModel, Field, ValidationError


# --------------------------
# Logger configuration block
# --------------------------

logger = logging.getLogger("models")
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
# Default level can be adjusted by the main app; keep INFO here
logger.setLevel(logging.INFO)


# --------------------------------
# Shared mixin for safe model logs
# --------------------------------

class LoggableModel(BaseModel):
    """
    Mixin that logs model creation and provides a safe_dump helper for debug.
    Uses Pydantic v2 BaseModel.
    """

    def model_post_init(self, __context: Any) -> None:
        # Lightweight, structured log for model creation
        try:
            # Keep logs concise: only top-level keys
            payload_keys = list(self.model_dump().keys())
            logger.debug(f"{self.__class__.__name__} created with fields: {payload_keys}")
        except Exception as e:
            logger.debug(f"{self.__class__.__name__} post-init logging failed: {e}")

    def safe_dump(self) -> str:
        try:
            return json.dumps(self.model_dump(mode="json"), ensure_ascii=False)
        except Exception as e:
            return f"<safe_dump_error:{e}>"


# -----------------------
# Worksheet payload pieces
# -----------------------

class CellWrite(LoggableModel):
    address: str = Field(..., description="A1-style cell address, e.g., 'A1'")
    value: Any = Field(None, description="Scalar value to write at address")


class FormulaWrite(LoggableModel):
    address: str = Field(..., description="A1-style cell address for the formula")
    formula: str = Field(..., description="Excel formula string beginning with '='")
    description: Optional[str] = Field(None, description="Optional human-readable description")


class ConditionalFormatHint(LoggableModel):
    range: str = Field(..., description="Target range, e.g., 'A1:C10' or 'D:D'")
    type: Optional[str] = Field(
        None,
        description="Frontend-specific type hint (e.g., 'colorScale', 'greaterThan')"
    )
    style: Optional[str] = Field(None, description="Frontend-specific style name")
    threshold: Optional[float] = Field(None, description="Optional numeric threshold")


class WorksheetPayload(LoggableModel):
    name: str = Field(..., description="Worksheet name to create or update")
    cells: List[CellWrite] = Field(default_factory=list)
    formulas: List[FormulaWrite] = Field(default_factory=list)
    notes: Optional[str] = Field(None, description="Optional notes to display on the sheet")
    conditional_formatting: Optional[List[ConditionalFormatHint]] = Field(
        default=None,
        description="Optional formatting hints for the frontend"
    )
    column_widths: Optional[List[Tuple[str, float]]] = Field(
        default=None,
        description="Optional list of (column_letter, width)"
    )
    freeze_panes: Optional[str] = Field(
        default=None,
        description="Optional freeze pane anchor, e.g., 'A2'"
    )


# -------------------------
# Core request/response IOs
# -------------------------

class TablePayload(LoggableModel):
    table: List[List[Any]] = Field(..., description="2D array from Excel selection")
    address: Optional[str] = Field(None, description="Excel address of selection, e.g., 'Sheet1!A1:C10'")

    @classmethod
    def from_raw(cls, data: Dict[str, Any]) -> "TablePayload":
        try:
            model = cls.model_validate(data)
            logger.debug(f"TablePayload validated: rows={len(model.table)} address={model.address}")
            return model
        except ValidationError as e:
            logger.error(f"TablePayload validation error: {e}")
            raise


class AnalyzeRequest(TablePayload):
    question: str = Field(..., description="Natural language question")


class AnalyzeResponse(LoggableModel):
    answer: str = Field(..., description="Concise analysis answer")


class AuditRow(LoggableModel):
    check: str = Field(..., description="Name of the audit check")
    result: str = Field(..., description="Result of the audit check")
    description: str = Field(default="", description="LLM-generated plain-English explanation")


class AuditResponse(LoggableModel):
    audit: str = Field(..., description="Data quality summary string (backward compat)")
    rows: List[AuditRow] = Field(default_factory=list, description="Structured rows for tabular display")

class IntentResponse(LoggableModel):
    action_type: str = Field(..., description="One of: analyze, audit, chart, unknown")
    confidence: float = Field(..., description="Classifier confidence between 0 and 1")
    data_complexity: str = Field(..., description="One of: small, medium, large")


class ActionResponse(LoggableModel):
    success: bool = Field(..., description="Overall success flag for the action")
    message: str = Field(..., description="Human-readable outcome message")
    worksheets: List[WorksheetPayload] = Field(default_factory=list, description="List of worksheet plans")
    formulas: List[Dict[str, Any]] = Field(
        default_factory=list,
        description="Deprecated/unused: formulas are embedded per worksheet"
    )

    def append_sheet(self, sheet: WorksheetPayload) -> None:
        try:
            self.worksheets.append(sheet)
            logger.debug(f"Appended worksheet: {sheet.name}")
        except Exception as e:
            logger.error(f"Failed to append worksheet: {e}")

class AiChartRequest(BaseModel):
    table: list
    address: Optional[str] = None
    question: Optional[str] = Field(default=None)
    use_llm: Optional[bool] = Field(default=False)

class ChartSuggestResponse(LoggableModel):
    chartType: str = Field(..., description="Excel chart type identifier (e.g., 'ColumnClustered', 'Line')")
    chartTitle: str = Field(..., description="Suggested chart title")
    inferences: str = Field(default="", description="Short analytical summary of the chart")
    categoryColumn: Optional[Union[str, List[str]]] = None
    valueColumns: List[Any] = []
    aggregations: List[str] = []


# -------------------------
# Smart Routing models
# -------------------------

class SmartRouteRequest(LoggableModel):
    table: List[List[Any]] = Field(..., description="2D array from Excel selection")
    address: Optional[str] = Field(None, description="Excel address of selection")
    question: str = Field(..., description="User's natural language question")
    context: Optional[str] = Field(None, description="Prior conversation turn / clarification answer")
    clarification_round: int = Field(default=0, description="How many HITL rounds have occurred (max 2)")


class SmartRouteResponse(LoggableModel):
    """
    Returned by /smart_route.

    When needs_clarification=True, the frontend should render the
    clarification_question and clarification_options and NOT yet execute anything.

    When needs_clarification=False, the frontend executes based on mode.
    """
    mode: str = Field(..., description="Chosen execution mode")
    confidence: float = Field(..., description="Router confidence 0.0–1.0")
    rationale: str = Field(default="", description="Why this mode was chosen (shown when 0.55–0.79)")
    needs_clarification: bool = Field(default=False)
    clarification_question: Optional[str] = None
    clarification_options: List[str] = Field(default_factory=list)
    # If mode resolved without clarification, downstream results are embedded:
    action_result: Optional[ActionResponse] = None
    chart_result: Optional[ChartSuggestResponse] = None
    analyze_result: Optional[AnalyzeResponse] = None