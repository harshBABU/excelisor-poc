# aibackend/intent.py

from __future__ import annotations

import json
import logging
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple, Union
from dataclasses import dataclass
from enum import Enum

import pandas as pd
from llm_config import get_openai_client, get_model

# Optional OpenAI integration
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

# Setup logging
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


class ActionType(Enum):
    """Enhanced action types for Excel assistant operations."""
    ANALYZE = "analyze"
    AUDIT = "audit" 
    CHART = "chart"
    FORMULA = "formula"
    FILTER = "filter"
    #PIVOT = "pivot"
    SORT = "sort"
    SUMMARY = "summary"
    COMPARE = "compare"
    FORECAST = "forecast"
    UNKNOWN = "unknown"


class DataComplexity(Enum):
    """Data complexity levels for processing decisions."""
    SIMPLE = "simple"      # < 100 cells
    MEDIUM = "medium"      # 100-5000 cells
    LARGE = "large"        # 5000-50000 cells
    MASSIVE = "massive"    # > 50000 cells


class ConfidenceLevel(Enum):
    """Confidence levels for intent classification."""
    VERY_LOW = "very_low"    # 0.0-0.3
    LOW = "low"              # 0.3-0.5
    MEDIUM = "medium"        # 0.5-0.7
    HIGH = "high"            # 0.7-0.9
    VERY_HIGH = "very_high"  # 0.9-1.0


@dataclass
class EntityExtraction:
    """Extracted entities from user query."""
    metrics: List[str]           # Numerical measures (sales, profit, quantity)
    dimensions: List[str]        # Categorical groupings (product, region, date)
    aggregations: List[str]      # Statistical operations (sum, avg, max)
    time_references: List[str]   # Temporal contexts (last month, Q1, yearly)
    comparisons: List[str]       # Comparative terms (vs, compared to, better than)
    thresholds: List[float]      # Numerical thresholds (> 1000, < 50%)
    operations: List[str]        # Specific operations (filter, sort, rank)


@dataclass
class IntentResult:
    """Comprehensive intent classification result."""
    action_type: ActionType
    confidence: float
    confidence_level: ConfidenceLevel
    data_complexity: DataComplexity
    entities: EntityExtraction
    reasoning: str
    suggested_columns: Dict[str, List[str]]  # {"numeric": [...], "categorical": [...]}
    analysis_plan: Optional[Dict[str, Any]] = None
    llm_enhanced: bool = False
    fallback_used: bool = False


class EnhancedIntentClassifier:
    """
    LLM-enhanced intent classifier for robust Excel assistant operations.
    Provides intelligent intent classification, entity extraction, and column selection.
    """
    
    def __init__(self):
        self.openai_client = self._initialize_openai_client()
        self.keyword_patterns = self._initialize_keyword_patterns()
        logger.info("Enhanced Intent Classifier initialized")
    
    def _initialize_openai_client(self) -> Optional[OpenAI]:
        """Initialize OpenAI client if available."""
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key or not OpenAI:
            logger.warning("OpenAI not available - using fallback classification")
            return None
        
        try:
            from llm_config import get_openai_client
            client = get_openai_client()
            logger.info("OpenAI client initialized successfully")
            return client
        except Exception as e:
            logger.error(f"Failed to initialize OpenAI client: {e}")
            return None
    
    def _initialize_keyword_patterns(self) -> Dict[ActionType, Dict[str, List[str]]]:
        """Initialize keyword patterns for fallback classification."""
        return {
            ActionType.ANALYZE: {
                "primary": ["analyze", "analysis", "insight", "understand", "explore", "examine"],
                "secondary": ["what", "how", "why", "trend", "pattern", "relationship"]
            },
            ActionType.AUDIT: {
                "primary": ["audit", "quality", "clean", "validate", "check", "missing"],
                "secondary": ["null", "duplicate", "error", "inconsistent", "invalid"]
            },
            ActionType.CHART: {
                "primary": ["chart", "plot", "graph", "visual", "visualize", "show"],
                "secondary": ["bar", "line", "pie", "scatter", "histogram", "dashboard"]
            },
            ActionType.FORMULA: {
                "primary": ["formula", "calculate", "compute", "derive", "equation"],
                "secondary": ["sum", "average", "vlookup", "if", "count", "index"]
            },
            ActionType.FILTER: {
                "primary": ["filter", "where", "select", "only", "exclude", "include"],
                "secondary": ["greater", "less", "equal", "between", "contains", "match"]
            },
            #ActionType.PIVOT: {
            #    "primary": ["pivot", "summarize", "group by", "aggregate", "breakdown"],
            #    "secondary": ["cross-tab", "summary table", "group", "subtotal"]
            #},
            ActionType.SORT: {
                "primary": ["sort", "order", "rank", "arrange", "sequence"],
                "secondary": ["ascending", "descending", "top", "bottom", "highest", "lowest"]
            },
            ActionType.SUMMARY: {
                "primary": ["summary", "overview", "total", "count", "statistics"],
                "secondary": ["stats", "describe", "distribution", "frequency"]
            },
            ActionType.COMPARE: {
                "primary": ["compare", "vs", "versus", "difference", "contrast"],
                "secondary": ["against", "between", "relative", "benchmark"]
            },
            ActionType.FORECAST: {
                "primary": ["forecast", "predict", "projection", "future", "trend"],
                "secondary": ["next", "upcoming", "estimate", "extrapolate"]
            }
        }
    
    def classify_intent(self, question: str, df: pd.DataFrame, context: Optional[str] = None) -> IntentResult:
        """
        Main entry point for intent classification with LLM enhancement.
        
        Args:
            question: User's natural language query
            df: DataFrame containing the data
            context: Optional conversation context
            
        Returns:
            IntentResult with comprehensive classification details
        """
        logger.info(f"Classifying intent for question: '{question[:100]}...'")
        
        # Data complexity assessment
        data_complexity = self._assess_data_complexity(df)
        
        # Try LLM-enhanced classification first
        if self.openai_client:
            try:
                result = self._llm_classify_intent(question, df, context)
                if result is not None:  # Check if LLM classification succeeded
                    result.data_complexity = data_complexity
                    logger.info(f"LLM classification successful: {result.action_type.value} (confidence: {result.confidence:.2f})")
                    return result
                else:
                    logger.warning("LLM classification returned None, using fallback")
            except Exception as e:
                logger.warning(f"LLM classification failed, using fallback: {str(e)}")
        
        # Fallback to rule-based classification
        result = self._fallback_classify_intent(question, df, context)
        result.data_complexity = data_complexity
        result.fallback_used = True
        logger.info(f"Fallback classification: {result.action_type.value} (confidence: {result.confidence:.2f})")
        return result
    
    def _llm_classify_intent(self, question: str, df: pd.DataFrame, context: Optional[str] = None) -> IntentResult:
        """LLM-powered intent classification with comprehensive analysis."""
        
        # Build data context for LLM
        data_context = self._build_data_context(df)
        
        # Create system prompt
        system_prompt = self._create_classification_system_prompt()
        
        # Create user prompt
        user_prompt = self._create_classification_user_prompt(question, data_context, context)
        
        # Call LLM with JSON format enforcement
        response = self.openai_client.chat.completions.create(
            model=get_model(),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=1,  # Lower temperature for more structured output
            response_format={"type": "json_object"}  # Enforce JSON output
            #max_tokens=1500
        )
        
        # Parse LLM response with error handling
        response_content = response.choices[0].message.content
        logger.debug(f"Raw LLM response content: {response_content[:500]}...")  # First 500 chars for debugging
        print(f"[DEBUG] Full LLM response: {response_content}")
        
        try:
            result_data = json.loads(response_content)
        except json.JSONDecodeError as e:
            logger.error(f"JSON parsing failed: {str(e)}")
            logger.error(f"Raw response content: {response_content}")
            # Return None to trigger fallback
            return None
        
        # Convert to IntentResult
        return self._parse_llm_response(result_data, df)
    
    def _create_classification_system_prompt(self) -> str:
        """Create system prompt for LLM intent classification with enhanced multi-dimensional query support.
        
        Recent fix: Enhanced multi-dimensional query detection and parsing with support for:
        - Multiple aggregation functions in one query ("sum X and average Y")
        - Multiple groupby dimensions ("by column1 and column2")
        - Complex patterns like "top 5 products by sales with average margin"
        - Improved pattern recognition with comprehensive examples
        - Enhanced JSON response structure for better multi-dimensional capture
        """
        action_types = [action.value for action in ActionType]
        
        return f"""You are an expert Excel data analyst that classifies user intents and extracts relevant information, with ADVANCED support for multi-dimensional queries and complex analytical patterns.

Your task is to analyze user questions about Excel data and return a structured JSON response with:

1. ACTION CLASSIFICATION: Choose from {action_types}
   - analyze: General data exploration and insights
   - audit: Data quality checks, missing values, duplicates
   - chart: Create visualizations (charts, graphs, plots)
   - formula: Create calculations or Excel formulas (PREFER for multi-aggregation queries)
   - filter: Filter/subset data based on conditions
   - sort: Sort or rank data
   - summary: Statistical summaries and descriptions
   - compare: Compare different groups or time periods
   - forecast: Predictive analysis or trend projection
   - unknown: Cannot determine intent

CRITICAL: Multi-dimensional queries with multiple aggregations should be classified as "formula":
- "sum qty and average revenue by sku" → formula (multi-aggregation)
- "count orders and total sales by region and category" → formula (multi-aggregation + multi-groupby)
- "top 5 products by sales with average margin" → formula (ranking + aggregation + complex filtering)

2. ENHANCED MULTI-DIMENSIONAL QUERY DETECTION: Identify complex patterns with high precision:

A. MULTIPLE AGGREGATIONS: Queries with multiple calculations
   Pattern indicators: 
   - Conjunctions: "and", ",", "plus", "along with", "together with"
   - Multiple function words: "sum X and average Y", "count A, total B and max C"
   - Comparative aggregations: "both sum and average", "min and max values"
   - Sequential patterns: "calculate sum, then average", "find total and mean"
   
   Enhanced Examples:
   - "sum qty and average price" → ["sum", "average"]
   - "count orders, total sales and max discount" → ["count", "sum", "max"]
   - "minimum cost and maximum profit by category" → ["min", "max"]
   - "total revenue plus average margin per product" → ["sum", "average"]
   - "both count and sum of orders by region" → ["count", "sum"]
   - "calculate mean price along with total quantity" → ["average", "sum"]
   - "show sum, average, and count for each group" → ["sum", "average", "count"]

B. MULTIPLE GROUPBY DIMENSIONS: Queries with multiple grouping columns
   Pattern indicators: 
   - Multi-dimensional grouping: "by X and Y", "by X, Y and Z", "grouped by multiple"
   - Hierarchical grouping: "by region then category", "first by A then by B"
   - Cross-tabulation patterns: "across X and Y", "breakdown by X and Y"
   - Nested grouping: "within each X, group by Y"
   
   Enhanced Examples:
   - "sales by region and category" → groupby: ["region", "category"]
   - "count by sku, date and status" → groupby: ["sku", "date", "status"] 
   - "average revenue by supplier and product type" → groupby: ["supplier", "product_type"]
   - "breakdown by region, category, and quarter" → groupby: ["region", "category", "quarter"]
   - "group by customer type and geography" → groupby: ["customer_type", "geography"]
   - "analyze across product lines and sales channels" → groupby: ["product_line", "sales_channel"]
   - "segment by age group, income bracket, and location" → groupby: ["age_group", "income_bracket", "location"]

C. COMPLEX RANKING WITH AGGREGATIONS: Multi-step analytical queries
   Pattern indicators: 
   - Ranking with secondary metrics: "top N ... with/by ...", "best/worst ... with average/total"
   - Conditional aggregations: "highest X where Y", "top performers with metrics"
   - Multi-criteria analysis: "best by X showing Y and Z", "rank by A with B and C"
   - Performance analysis: "leading products by sales with profitability metrics"
   
   Enhanced Examples:
   - "top 5 products by sales with average margin" → ranking + secondary aggregation
   - "worst 3 regions by revenue with total orders" → ranking + secondary metric
   - "best performing categories with sum of profit and count of products" → ranking + multi-aggregation
   - "highest revenue products showing average price and total quantity" → ranking + multi-aggregation
   - "top 10 customers by spend with average order value and frequency" → ranking + multi-metrics
   - "bottom 5 suppliers by quality score with total orders and average delivery time" → ranking + multi-aggregation

3. ENTITY EXTRACTION: Identify specific data elements mentioned:
   - metrics: Numerical measures (sales, revenue, quantity, profit, cost, margin)
   - dimensions: Categorical groupings (product, region, customer, date, sku, category, supplier)
   - aggregations: Statistical operations (sum, average, max, min, count, median, std, total)
   - time_references: Temporal contexts (last month, Q1, yearly, trend)
   - comparisons: Comparative terms (vs, compared to, better than)
   - thresholds: Numerical thresholds (> 1000, top 10, bottom 5, < 50%)
   - operations: Specific operations (filter, sort, group, rank, top, bottom)

ENHANCED AGGREGATION DETECTION RULES:
- Multiple aggregations: Look for "and", ",", "plus", "along with", "together with", "as well as"
- Function synonyms: 
  * sum/total/aggregate/summation
  * average/mean/avg/typical
  * count/frequency/number of/occurrences
  * max/maximum/highest/largest/peak
  * min/minimum/lowest/smallest/bottom
  * median/middle/50th percentile
  * std/stddev/standard deviation/variance
- Hidden aggregations: "with average", "showing total", "including count", "plus mean"
- Implicit calculations: "margin" (profit/cost), "ratio" (division), "percentage" (proportion)
- Compound patterns: "sum and average of", "total plus count of", "min, max and average"
- Context-aware detection: "revenue (total and average)", "sales: sum and mean"

CRITICAL COLUMN IDENTIFICATION RULES:
- When the user asks to group "by SKU" or mentions "sku", ALWAYS use the SKU column, NOT the ID column
- PRIORITY ORDER for SKU-related queries: sku_code > sku > product_code > product_id > item_id > id (LAST RESORT)
- ID columns are typically UUIDs and should NOT be used for grouping in business analysis
- SKU columns contain business-meaningful product identifiers and should be preferred for grouping
- Look for column names containing "sku_code", "sku", "product_code", "product", "item" for groupby operations
- Avoid using columns with names like "id", "uuid", "guid" for grouping unless explicitly requested
- For multi-groupby: prioritize business columns over technical columns
- NEVER choose "id" when "sku_code" is available for SKU-related operations

COMPREHENSIVE MULTI-DIMENSIONAL EXECUTION PLANS:

Example 1: "sum qty and average revenue by sku and region"
- is_multi_dimensional: true
- aggregations: ["sum", "average"], metrics: ["qty", "revenue"], dimensions: ["sku", "region"]
- multi_aggregation: true, multi_groupby: true
- aggregation_details: [{{"function": "sum", "column": "qty"}}, {{"function": "average", "column": "revenue"}}]
- groupby_columns: ["SKU", "region"]
- complexity_level: "medium"
- execution_steps: ["Group data by SKU and region", "Calculate sum of qty for each group", "Calculate average revenue for each group", "Generate multi-criteria Excel formulas"]

Example 2: "top 5 products by sales with average margin and count of orders"
- is_multi_dimensional: true
- aggregations: ["sum", "average", "count"], thresholds: [5], operations: ["top", "rank"]
- python_preprocessing: true, ranking_criteria: "sales", filter_operation: "top_n", filter_value: 5
- multi_aggregation: true (secondary aggregations after ranking)
- complexity_level: "complex"
- aggregation_details: [{{"function": "average", "column": "margin"}}, {{"function": "count", "column": "orders"}}]
- execution_steps: ["Rank products by sales", "Filter to top 5", "Calculate average margin for each", "Count orders for each", "Generate multi-aggregation results"]

Example 3: "count orders and total sales by region and category where status = 'active'"
- is_multi_dimensional: true
- aggregations: ["count", "sum"], dimensions: ["region", "category"], operations: ["filter"]
- multi_aggregation: true, multi_groupby: true
- complexity_level: "medium"
- filter_conditions: [{{"column": "status", "operator": "=", "value": "active"}}]
- aggregation_details: [{{"function": "count", "column": "orders"}}, {{"function": "sum", "column": "sales"}}]
- groupby_columns: ["region", "category"]
- has_filters: true

Example 4: "sum, average, and count of revenue by product line, region, and quarter"
- is_multi_dimensional: true
- aggregations: ["sum", "average", "count"], metrics: ["revenue"], dimensions: ["product_line", "region", "quarter"]
- multi_aggregation: true, multi_groupby: true
- complexity_level: "complex"
- aggregation_details: [{{"function": "sum", "column": "revenue"}}, {{"function": "average", "column": "revenue"}}, {{"function": "count", "column": "revenue"}}]
- groupby_columns: ["product_line", "region", "quarter"]
- execution_steps: ["Group data by product line, region, and quarter", "Calculate sum of revenue for each group", "Calculate average revenue for each group", "Count revenue entries for each group", "Generate comprehensive multi-aggregation analysis"]

Example 5: "highest margin products showing total quantity and average price"
- is_multi_dimensional: true
- aggregations: ["max", "sum", "average"], operations: ["rank", "filter"]
- multi_aggregation: true
- complexity_level: "complex"
- ranking_criteria: "margin", filter_operation: "max"
- aggregation_details: [{{"function": "sum", "column": "quantity"}}, {{"function": "average", "column": "price"}}]
- execution_steps: ["Identify products with highest margin", "Calculate total quantity for high-margin products", "Calculate average price for high-margin products", "Generate ranking analysis with secondary metrics"]

Example 6: "total revenue plus average order value by customer type and geography"
- is_multi_dimensional: true
- aggregations: ["sum", "average"], metrics: ["revenue", "order_value"], dimensions: ["customer_type", "geography"]
- multi_aggregation: true, multi_groupby: true
- complexity_level: "medium"
- aggregation_details: [{{"function": "sum", "column": "revenue"}}, {{"function": "average", "column": "order_value"}}]
- groupby_columns: ["customer_type", "geography"]
- execution_steps: ["Group data by customer type and geography", "Calculate total revenue for each group", "Calculate average order value for each group", "Combine multi-aggregation results"]

Example 7: "average revenue by sku of top 5 sku by sales" (Legacy Complex)
- aggregations: ["average"], thresholds: [5], operations: ["top", "rank"]
- python_preprocessing: true, ranking_criteria: "sales", filter_operation: "top_n", filter_value: 5
- complexity_level: "medium"
- final_aggregation: "average", aggregation_column: "revenue", groupby_column: "SKU"
- execution_steps: ["Rank all SKUs by sales amount", "Filter to top 5 SKUs", "Calculate average revenue by SKU for filtered data"]

ADDITIONAL COMPREHENSIVE EXAMPLES:

Example 8: "sum sales and count orders by region and month"
- pattern_type: "combined", query_intent_pattern: "aggregate_and_group"
- aggregation_complexity: "multiple_different_columns", grouping_complexity: "multiple_independent"
- aggregation_details: [{{"function": "sum", "column": "sales", "alias": "total_sales"}}, {{"function": "count", "column": "orders", "alias": "order_count"}}]
- groupby_columns: ["region", "month"], complexity_level: "medium"

Example 9: "show sum, average, and count for revenue grouped by category"
- pattern_type: "multi_agg", query_intent_pattern: "aggregate_and_group"
- aggregation_complexity: "multiple_same_column", grouping_complexity: "single"
- aggregation_details: [{{"function": "sum", "column": "revenue", "alias": "total_revenue"}}, {{"function": "average", "column": "revenue", "alias": "avg_revenue"}}, {{"function": "count", "column": "revenue", "alias": "revenue_count"}}]
- primary_metric: "revenue", groupby_columns: ["category"], complexity_level: "medium"

Example 10: "total revenue plus average margin across product lines and regions"
- pattern_type: "combined", query_intent_pattern: "compare_across_dimensions"
- aggregation_complexity: "multiple_different_columns", grouping_complexity: "cross_tabulation"
- aggregation_details: [{{"function": "sum", "column": "revenue", "alias": "total_revenue"}}, {{"function": "average", "column": "margin", "alias": "avg_margin"}}]
- requires_pivoting: true, groupby_columns: ["product_line", "region"], complexity_level: "medium"

Example 11: "best 3 products by profit showing total quantity and average price"
- pattern_type: "complex_ranking", query_intent_pattern: "rank_with_metrics"
- aggregation_complexity: "multiple_different_columns", ranking_criteria: "profit", ranking_direction: "desc"
- filter_operation: "top_n", filter_value: 3, python_preprocessing: true
- aggregation_details: [{{"function": "sum", "column": "quantity", "alias": "total_qty"}}, {{"function": "average", "column": "price", "alias": "avg_price"}}]
- complexity_level: "complex"

4. COLUMN SELECTION: Based on the question and available columns, suggest the most relevant numeric and categorical columns.

5. CONFIDENCE ASSESSMENT: Rate your confidence (0.0-1.0) based on:
   - Clarity of the question
   - Specificity of requirements
   - Availability of relevant data columns
   - Complexity of multi-dimensional requirements

6. ENHANCED ANALYSIS PLAN: For multi-dimensional and complex queries:

REQUIRED FIELDS FOR ALL QUERIES:
   - multi_step_analysis: true/false (set to true for queries requiring multiple operations)
   - execution_steps: ["step1", "step2", "step3"] (ordered list of operations needed)
   - python_preprocessing: true/false (whether Python filtering/ranking is needed before Excel)

MULTI-DIMENSIONAL DETECTION FIELDS:
   - is_multi_dimensional: true/false (true if multiple aggregations OR multiple groupby columns OR complex ranking)
   - multi_aggregation: true/false (true if multiple aggregation functions requested)
   - multi_groupby: true/false (true if multiple groupby columns requested)
   - complexity_level: "simple" | "medium" | "complex"
   - pattern_type: "multi_agg" | "multi_groupby" | "complex_ranking" | "combined" | "simple"

ENHANCED AGGREGATION SPECIFICATION FIELDS:
   - aggregation_details: [{{"function": "sum", "column": "qty", "alias": "total_qty"}}, {{"function": "average", "column": "revenue", "alias": "avg_revenue"}}]
   - groupby_columns: ["column1", "column2"] (list of all groupby columns for multi-dimensional queries)
   - primary_metric: "column_name" (main metric being analyzed)
   - secondary_metrics: ["metric1", "metric2"] (additional metrics in multi-aggregation queries)
   - ranking_criteria: "column_name" (which column to use for ranking, if applicable)
   - ranking_direction: "desc" | "asc" (direction for ranking operations)
   - filter_operation: "top_n" | "bottom_n" | "threshold" | "percentile" | "conditional" (type of filtering needed)
   - filter_value: number (N for top/bottom N, or threshold value)

ADVANCED PATTERN RECOGNITION FIELDS:
   - query_intent_pattern: "aggregate_and_group" | "rank_with_metrics" | "compare_across_dimensions" | "trend_analysis"
   - aggregation_complexity: "single" | "multiple_same_column" | "multiple_different_columns" | "nested_calculations"
   - grouping_complexity: "single" | "multiple_independent" | "hierarchical" | "cross_tabulation"
   - requires_pivoting: true/false (whether query benefits from pivot table structure)
   - temporal_analysis: true/false (whether query involves time-based analysis)

LEGACY COMPATIBILITY FIELDS (for single aggregation queries):
   - final_aggregation: "sum" | "average" | "max" | "min" | "count" (final Excel operation)
   - aggregation_column: "column_name" (which column to aggregate)
   - groupby_column: "column_name" (which column to group by for final aggregation)

ENHANCED FILTER SPECIFICATION FIELDS:
   - filter_conditions: [{{"column": "status", "operator": "=", "value": "active", "logic": "AND", "data_type": "categorical"}}]
   - has_filters: true/false
   - filter_complexity: "simple" | "multiple_conditions" | "nested_logic" | "dynamic_thresholds"

Return ONLY valid JSON in this exact format:
{{
    "action_type": "string",
    "confidence": 0.0-1.0,
    "reasoning": "explanation of classification decision and multi-dimensional detection",
    "entities": {{
        "metrics": ["list", "of", "metrics"],
        "dimensions": ["list", "of", "dimensions"],
        "aggregations": ["list", "of", "aggregations"],
        "time_references": ["list", "of", "time_refs"],
        "comparisons": ["list", "of", "comparisons"],
        "thresholds": [1000, 50.5],
        "operations": ["list", "of", "operations"]
    }},
    "suggested_columns": {{
        "numeric": ["column1", "column2"],
        "categorical": ["column3", "column4"]
    }},
    "analysis_plan": {{
        "primary_focus": "string",
        "secondary_insights": ["list", "of", "insights"],
        "recommended_charts": ["chart", "types"],
        "complexity_factors": ["factors", "to", "consider"],
        "multi_step_analysis": true/false,
        "execution_steps": ["step1", "step2", "step3"],
        "python_preprocessing": true/false,
        "is_multi_dimensional": true/false,
        "multi_aggregation": true/false,
        "multi_groupby": true/false,
        "complexity_level": "simple|medium|complex",
        "pattern_type": "multi_agg|multi_groupby|complex_ranking|combined|simple",
        "aggregation_details": [{{"function": "string", "column": "string", "alias": "string"}}],
        "groupby_columns": ["column1", "column2"],
        "primary_metric": "column_name_or_null",
        "secondary_metrics": ["metric1", "metric2"],
        "ranking_criteria": "column_name_or_null",
        "ranking_direction": "desc|asc|null",
        "filter_operation": "top_n|bottom_n|threshold|percentile|conditional|null",
        "filter_value": number_or_null,
        "query_intent_pattern": "aggregate_and_group|rank_with_metrics|compare_across_dimensions|trend_analysis|null",
        "aggregation_complexity": "single|multiple_same_column|multiple_different_columns|nested_calculations",
        "grouping_complexity": "single|multiple_independent|hierarchical|cross_tabulation",
        "requires_pivoting": true/false,
        "temporal_analysis": true/false,
        "final_aggregation": "aggregation_function_or_null",
        "aggregation_column": "column_name_or_null",
        "groupby_column": "column_name_or_null",
        "filter_conditions": [{{"column": "string", "operator": "string", "value": "string", "logic": "AND|OR", "data_type": "string"}}],
        "has_filters": true/false,
        "filter_complexity": "simple|multiple_conditions|nested_logic|dynamic_thresholds"
    }}
}}"""
    
    def _create_classification_user_prompt(self, question: str, data_context: Dict[str, Any], context: Optional[str] = None) -> str:
        """Create user prompt with question and data context."""
        context_section = f"\n\nCONVERSATION CONTEXT:\n{context}" if context else ""
        
        return f"""QUESTION: {question}

DATA CONTEXT:
- Shape: {data_context['shape']}
- Columns: {data_context['columns']}
- Numeric columns: {data_context['numeric_columns']}
- Categorical columns: {data_context['categorical_columns']}
- Sample data preview:
{data_context['sample_data']}{context_section}

Please classify this intent and extract relevant entities. Focus on understanding what the user wants to accomplish with their Excel data."""
    
    def _build_data_context(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Build comprehensive data context for LLM analysis."""
        try:
            # Basic shape information
            shape = f"{df.shape[0]} rows, {df.shape[1]} columns"
            columns = list(df.columns)
            
            # Identify column types
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
            
            # Create sample data preview (first 3 rows, limited columns)
            preview_df = df.head(3)
            if len(columns) > 8:
                # Limit to first 8 columns for readability
                preview_df = preview_df.iloc[:, :8]
                columns_note = f" (showing first 8 of {len(columns)} columns)"
            else:
                columns_note = ""
            
            sample_data = preview_df.to_string(index=False, max_cols=8) + columns_note
            
            return {
                "shape": shape,
                "columns": columns,
                "numeric_columns": numeric_cols,
                "categorical_columns": categorical_cols,
                "sample_data": sample_data
            }
        except Exception as e:
            logger.error(f"Error building data context: {e}")
            return {
                "shape": "unknown",
                "columns": [],
                "numeric_columns": [],
                "categorical_columns": [],
                "sample_data": "Error loading data preview"
            }
    
    def _parse_llm_response(self, response_data: Dict[str, Any], df: pd.DataFrame) -> IntentResult:
        """Parse LLM response into IntentResult object."""
        try:
            # Parse action type
            action_type = ActionType(response_data.get("action_type", "unknown"))
            
            # Parse confidence and determine level
            confidence = float(response_data.get("confidence", 0.5))
            confidence_level = self._get_confidence_level(confidence)
            
            # Parse entities
            entities_data = response_data.get("entities", {})
            entities = EntityExtraction(
                metrics=entities_data.get("metrics", []),
                dimensions=entities_data.get("dimensions", []),
                aggregations=entities_data.get("aggregations", []),
                time_references=entities_data.get("time_references", []),
                comparisons=entities_data.get("comparisons", []),
                thresholds=entities_data.get("thresholds", []),
                operations=entities_data.get("operations", [])
            )
            
            # Validate and enhance column suggestions
            suggested_columns = self._validate_column_suggestions(
                response_data.get("suggested_columns", {}), df
            )
            
            # Recent fix: Validate analysis plan column references against actual DataFrame columns
            analysis_plan = response_data.get("analysis_plan")
            if analysis_plan:
                print(f"[DEBUG] Raw analysis_plan from LLM: {analysis_plan}")
                print(f"[DEBUG] Available DataFrame columns: {list(df.columns)}")
                analysis_plan = self._validate_analysis_plan_columns(analysis_plan, df)
                print(f"[DEBUG] Validated analysis_plan: {analysis_plan}")
            
            return IntentResult(
                action_type=action_type,
                confidence=confidence,
                confidence_level=confidence_level,
                data_complexity=DataComplexity.MEDIUM,  # Will be set by caller
                entities=entities,
                reasoning=response_data.get("reasoning", "LLM classification"),
                suggested_columns=suggested_columns,
                analysis_plan=analysis_plan,
                llm_enhanced=True,
                fallback_used=False
            )
            
        except Exception as e:
            logger.error(f"Error parsing LLM response: {e}")
            # Return fallback result
            return self._create_fallback_result("Error parsing LLM response", df)
    
    def _fallback_classify_intent(self, question: str, df: pd.DataFrame, context: Optional[str] = None) -> IntentResult:
        """Rule-based fallback classification when LLM is unavailable."""
        q = (question or "").lower().strip()
        
        # Extract entities first to help with classification
        entities = self._extract_entities_fallback(q)
        
        # Score each action type based on keyword matching
        action_scores = {}
        for action_type, patterns in self.keyword_patterns.items():
            score = 0
            
            # Primary keywords (higher weight)
            for keyword in patterns["primary"]:
                if keyword in q:
                    score += 2
            
            # Secondary keywords (lower weight)
            for keyword in patterns["secondary"]:
                if keyword in q:
                    score += 1
            
            action_scores[action_type] = score
        
        # Special handling for complex aggregation patterns
        # If we detect aggregation + ranking/filtering, it's likely a pivot operation
        has_aggregation = len(entities.aggregations) > 0
        has_ranking = any(op in entities.operations for op in ["top", "bottom", "rank"])
        has_grouping = any(dim in entities.dimensions for dim in ["sku", "product", "customer", "region"])
        has_threshold = len(entities.thresholds) > 0
        
        # Complex pattern: "average of sale for top 10 sku" should be pivot
        complex_agg_pattern = re.search(r'(average|sum|total|max|min|count|mean)\s+of\s+\w+\s+for\s+(top|bottom)\s+\d+', q)
        if complex_agg_pattern or (has_aggregation and has_ranking and (has_grouping or has_threshold)):
            action_scores[ActionType.PIVOT] += 4  # Strong boost for pivot classification
            logger.debug(f"Detected complex aggregation pattern, boosting pivot score")
        
        # Select best action type
        best_action = max(action_scores, key=action_scores.get)
        best_score = action_scores[best_action]
        
        if best_score == 0:
            best_action = ActionType.UNKNOWN
            confidence = 0.3
        else:
            # Normalize confidence based on score and question clarity
            max_possible_score = len(self.keyword_patterns[best_action]["primary"]) * 2 + \
                               len(self.keyword_patterns[best_action]["secondary"]) * 1
            confidence = min(0.9, 0.5 + (best_score / max_possible_score) * 0.4)
            
            # Boost confidence for detected complex patterns
            if complex_agg_pattern and best_action == ActionType.PIVOT:
                confidence = min(0.95, confidence + 0.2)
        
        # Select relevant columns using enhanced logic
        suggested_columns = self._select_relevant_columns_fallback(q, df)
        
        # Enhanced reasoning
        reasoning_parts = [f"Rule-based classification using keyword patterns (score: {best_score})"]
        if complex_agg_pattern:
            reasoning_parts.append("Detected complex aggregation pattern requiring pivot analysis")
        if has_aggregation:
            reasoning_parts.append(f"Found aggregations: {entities.aggregations}")
        if has_ranking:
            reasoning_parts.append(f"Found ranking operations: {[op for op in entities.operations if op in ['top', 'bottom', 'rank']]}")
        
        return IntentResult(
            action_type=best_action,
            confidence=confidence,
            confidence_level=self._get_confidence_level(confidence),
            data_complexity=DataComplexity.MEDIUM,  # Will be set by caller
            entities=entities,
            reasoning="; ".join(reasoning_parts),
            suggested_columns=suggested_columns,
            analysis_plan=None,
            llm_enhanced=False,
            fallback_used=True
        )
    
    def _extract_entities_fallback(self, question: str) -> EntityExtraction:
        """Extract entities using pattern matching as fallback."""
        q = question.lower()
        
        # Enhanced aggregation pattern detection
        aggregations = []
        agg_patterns = [
            ("sum", "SUM"), ("total", "SUM"), ("average", "AVERAGE"), ("avg", "AVERAGE"), 
            ("mean", "AVERAGE"), ("max", "MAX"), ("maximum", "MAX"), ("min", "MIN"), 
            ("minimum", "MIN"), ("count", "COUNT"), ("median", "MEDIAN"), ("std", "STDEV")
        ]
        for pattern, normalized in agg_patterns:
            if pattern in q:
                aggregations.append(normalized)
        
        # Enhanced operation patterns including ranking with aggregation
        operations = []
        op_patterns = ["filter", "sort", "group", "rank", "top", "bottom"]
        for pattern in op_patterns:
            if pattern in q:
                operations.append(pattern)
        
        # Special pattern: "X of Y for top/bottom N Z" indicates ranking + aggregation
        # E.g., "average of sale for top 10 sku" or "sum of revenue for bottom 5 products"
        ranking_agg_pattern = re.search(r'(average|sum|total|max|min|count|mean)\s+of\s+\w+\s+for\s+(top|bottom)\s+(\d+)', q)
        if ranking_agg_pattern:
            agg_func = ranking_agg_pattern.group(1)
            rank_type = ranking_agg_pattern.group(2)  # top/bottom
            rank_num = ranking_agg_pattern.group(3)   # number
            
            # Ensure we capture the aggregation
            normalized_agg = "AVERAGE" if agg_func in ["average", "mean"] else agg_func.upper()
            if normalized_agg not in aggregations:
                aggregations.append(normalized_agg)
            
            # Ensure we capture the ranking operation
            if rank_type not in operations:
                operations.append(rank_type)
            if "rank" not in operations:
                operations.append("rank")
        
        # Time reference patterns
        time_refs = []
        time_patterns = ["month", "year", "quarter", "week", "day", "q1", "q2", "q3", "q4", "ytd", "mtd"]
        for pattern in time_patterns:
            if pattern in q:
                time_refs.append(pattern)
        
        # Comparison patterns
        comparisons = []
        comp_patterns = ["vs", "versus", "compared to", "against", "better than", "worse than"]
        for pattern in comp_patterns:
            if pattern in q:
                comparisons.append(pattern)
        
        # Enhanced threshold patterns
        thresholds = []
        
        # Extract numbers following "top" or "bottom"
        top_bottom_matches = re.findall(r'(?:top|bottom)\s+(\d+)', q)
        for match in top_bottom_matches:
            thresholds.append(float(match))
        
        # Extract other numeric thresholds
        other_threshold_matches = re.findall(r'[><=]\s*[\d,]+\.?\d*|[\d,]+\.?\d*\s*[%]?', q)
        for match in other_threshold_matches:
            nums = re.findall(r'[\d,]+\.?\d*', match)
            for num in nums:
                try:
                    val = float(num.replace(',', ''))
                    if val not in thresholds:  # Avoid duplicates
                        thresholds.append(val)
                except ValueError:
                    continue
        
        # Basic metric extraction (common business terms)
        metrics = []
        metric_patterns = ["sales", "sale", "revenue", "profit", "cost", "price", "amount", "value", "quantity", "qty"]
        for pattern in metric_patterns:
            if pattern in q:
                metrics.append(pattern)
        
        # Basic dimension extraction (common grouping terms)
        dimensions = []
        dimension_patterns = ["sku", "product", "customer", "region", "category", "type", "brand", "segment"]
        for pattern in dimension_patterns:
            if pattern in q:
                dimensions.append(pattern)
        
        return EntityExtraction(
            metrics=list(set(metrics)),  # Remove duplicates
            dimensions=list(set(dimensions)),  # Remove duplicates
            aggregations=list(set(aggregations)),  # Remove duplicates
            time_references=time_refs,
            comparisons=comparisons,
            thresholds=thresholds,
            operations=list(set(operations))  # Remove duplicates
        )
    
    def _select_relevant_columns_fallback(self, question: str, df: pd.DataFrame) -> Dict[str, List[str]]:
        """Enhanced column selection using semantic patterns."""
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        selected_numeric = []
        selected_categorical = []
        
        # Financial/Business keywords
        financial_keywords = ["sales", "revenue", "price", "amount", "cost", "profit", "value", "total"]
        for col in numeric_cols:
            if any(keyword in col.lower() for keyword in financial_keywords):
                selected_numeric.append(col)
        
        # Quantity keywords
        quantity_keywords = ["qty", "quantity", "count", "units", "stock", "inventory"]
        for col in numeric_cols:
            if any(keyword in col.lower() for keyword in quantity_keywords):
                selected_numeric.append(col)
        
        # Category keywords for grouping
        category_keywords = ["category", "type", "sku", "product", "customer", "region", "group", "class"]
        for col in categorical_cols:
            if any(keyword in col.lower() for keyword in category_keywords):
                selected_categorical.append(col)
        
        # Date/Time keywords
        date_keywords = ["date", "time", "month", "year", "period"]
        for col in categorical_cols:
            if any(keyword in col.lower() for keyword in date_keywords):
                selected_categorical.append(col)
        
        # If no specific selections, use intelligent defaults
        if not selected_numeric:
            selected_numeric = numeric_cols[:3]  # Top 3 numeric columns
        if not selected_categorical:
            selected_categorical = categorical_cols[:2]  # Top 2 categorical columns
        
        return {
            "numeric": list(dict.fromkeys(selected_numeric))[:5],  # Remove duplicates, limit to 5
            "categorical": list(dict.fromkeys(selected_categorical))[:3]  # Remove duplicates, limit to 3
        }
    
    def _assess_data_complexity(self, df: pd.DataFrame) -> DataComplexity:
        """Assess data complexity based on DataFrame characteristics."""
        try:
            total_cells = df.shape[0] * df.shape[1]
            
            if total_cells < 100:
                return DataComplexity.SIMPLE
            elif total_cells < 5000:
                return DataComplexity.MEDIUM
            elif total_cells < 50000:
                return DataComplexity.LARGE
            else:
                return DataComplexity.MASSIVE
        except Exception:
            return DataComplexity.MEDIUM
    
    def _get_confidence_level(self, confidence: float) -> ConfidenceLevel:
        """Convert numeric confidence to confidence level enum."""
        if confidence < 0.3:
            return ConfidenceLevel.VERY_LOW
        elif confidence < 0.5:
            return ConfidenceLevel.LOW
        elif confidence < 0.7:
            return ConfidenceLevel.MEDIUM
        elif confidence < 0.9:
            return ConfidenceLevel.HIGH
        else:
            return ConfidenceLevel.VERY_HIGH
    
    def _validate_column_suggestions(self, suggestions: Dict[str, List[str]], df: pd.DataFrame) -> Dict[str, List[str]]:
        """Validate and filter column suggestions against actual DataFrame columns."""
        available_cols = set(df.columns)
        
        validated = {
            "numeric": [col for col in suggestions.get("numeric", []) if col in available_cols],
            "categorical": [col for col in suggestions.get("categorical", []) if col in available_cols]
        }
        
        # If LLM suggestions are empty or invalid, use fallback
        if not validated["numeric"] and not validated["categorical"]:
            return self._select_relevant_columns_fallback("", df)
        
        return validated
    
    def _validate_analysis_plan_columns(self, analysis_plan: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate and fix column references in analysis plan against actual DataFrame columns.
        
        Recent fix: The LLM might suggest semantic column names like 'sku' or 'revenue' 
        but the actual columns might be named 'id', 'amount', etc. This function maps 
        LLM semantic names to actual column names.
        """
        if not analysis_plan:
            return analysis_plan
        
        validated_plan = analysis_plan.copy()
        available_cols = list(df.columns)
        available_cols_lower = [col.lower() for col in available_cols]
        
        # Column mappings that need validation
        column_mappings = {
            'aggregation_column': 'numeric',
            'groupby_column': 'categorical', 
            'ranking_criteria': 'numeric'
        }
        
        for plan_key, expected_type in column_mappings.items():
            if plan_key in validated_plan:
                suggested_col = validated_plan[plan_key]
                
                # Recent fix: Handle None values from LLM response - but only warn if it's not a multi-aggregation format
                if suggested_col is None or not isinstance(suggested_col, str):
                    # Check if this is a multi-aggregation format (new style) - if so, don't warn about missing old-style fields
                    is_multi_agg = (validated_plan.get('multi_aggregation') is True or 
                                   'aggregation_details' in validated_plan or 
                                   'groupby_columns' in validated_plan)
                    
                    if not is_multi_agg:
                        logger.warning(f"Analysis plan column '{plan_key}' is None or invalid type, removing")
                    validated_plan.pop(plan_key, None)
                    continue
                
                # If column exists exactly, keep it
                if suggested_col in available_cols:
                    logger.debug(f"Analysis plan column '{suggested_col}' found exactly in DataFrame")
                    continue
                
                # Try to find best match by semantic similarity
                best_match = self._find_best_column_match(suggested_col, available_cols, df, expected_type)
                
                if best_match and best_match != suggested_col:
                    logger.info(f"Analysis plan column '{suggested_col}' mapped to actual column '{best_match}'")
                    validated_plan[plan_key] = best_match
                elif not best_match:
                    logger.warning(f"Analysis plan column '{suggested_col}' could not be mapped to any DataFrame column")
                    # Remove the invalid column reference
                    validated_plan.pop(plan_key, None)
        
        # Recent fix: Validate new multi-aggregation format fields
        if 'groupby_columns' in validated_plan:
            groupby_columns = validated_plan['groupby_columns']
            if isinstance(groupby_columns, list):
                validated_groupby = []
                for col in groupby_columns:
                    if col in available_cols:
                        validated_groupby.append(col)
                    else:
                        best_match = self._find_best_column_match(col, available_cols, df, 'categorical')
                        if best_match:
                            logger.info(f"Multi-agg groupby column '{col}' mapped to '{best_match}'")
                            validated_groupby.append(best_match)
                        else:
                            logger.warning(f"Multi-agg groupby column '{col}' not found in DataFrame")
                validated_plan['groupby_columns'] = validated_groupby
        
        if 'aggregation_details' in validated_plan:
            agg_details = validated_plan['aggregation_details']
            if isinstance(agg_details, list):
                validated_agg_details = []
                for detail in agg_details:
                    if isinstance(detail, dict) and 'column' in detail:
                        col = detail['column']
                        if col in available_cols:
                            validated_agg_details.append(detail)
                        else:
                            best_match = self._find_best_column_match(col, available_cols, df, 'numeric')
                            if best_match:
                                logger.info(f"Multi-agg column '{col}' mapped to '{best_match}'")
                                detail_copy = detail.copy()
                                detail_copy['column'] = best_match
                                validated_agg_details.append(detail_copy)
                            else:
                                logger.warning(f"Multi-agg column '{col}' not found in DataFrame")
                    else:
                        validated_agg_details.append(detail)
                validated_plan['aggregation_details'] = validated_agg_details
        
        return validated_plan
    
    def _find_best_column_match(self, suggested_col: str, available_cols: List[str], df: pd.DataFrame, expected_type: str) -> Optional[str]:
        """Find the best matching column for LLM suggestions."""
        # Recent fix: Additional safety check for None values
        if not suggested_col or not isinstance(suggested_col, str):
            logger.warning(f"Invalid suggested_col: {suggested_col}")
            return None
            
        suggested_lower = suggested_col.lower()
        logger.debug(f"Looking for column match: '{suggested_col}' (type: {expected_type}) among {len(available_cols)} columns")
        
        # Get columns of the expected type
        if expected_type == 'numeric':
            typed_cols = df.select_dtypes(include=['number']).columns.tolist()
        else:
            typed_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        # First priority: Exact semantic match within typed columns
        # Recent fix: Prioritize exact matches and avoid short substring matches
        for col in typed_cols:
            col_lower = col.lower()
            # Exact match has highest priority
            if suggested_lower == col_lower:
                logger.debug(f"Found exact match: '{suggested_col}' -> '{col}'")
                return col
            # Longer meaningful substring matches (avoid single character matches)
            if len(suggested_lower) > 2 and suggested_lower in col_lower:
                logger.debug(f"Found substring match: '{suggested_col}' in '{col}'")
                return col
            if len(col_lower) > 2 and col_lower in suggested_lower:
                logger.debug(f"Found reverse substring match: '{col}' in '{suggested_col}'")
                return col
        
        # Second priority: Common business term mappings
        # Recent fix: Prioritize descriptive columns over generic IDs
        business_mappings = {
            'sku': ['sku_code', 'sku', 'code', 'product_code', 'item_code', 'product_id', 'item_id', 'id'],
            'revenue': ['revenue', 'sales', 'amount', 'total', 'value', 'price'],
            'sales': ['sales', 'revenue', 'amount', 'total', 'value'],
            'customer': ['customer', 'client', 'user', 'buyer', 'account'],
            'product': ['product', 'item', 'sku', 'good'],
            'quantity': ['qty', 'quantity', 'count', 'units', 'volume']
        }
        
        if suggested_lower in business_mappings:
            for mapping in business_mappings[suggested_lower]:
                for col in typed_cols:
                    col_lower = col.lower()
                    # Prioritize exact matches in business mapping
                    if mapping == col_lower:
                        return col
            # If no exact match, look for substring matches
            for mapping in business_mappings[suggested_lower]:
                for col in typed_cols:
                    col_lower = col.lower()
                    if mapping in col_lower:
                        return col
        
        # Third priority: Reverse mapping (column name suggests LLM term)
        for term, mappings in business_mappings.items():
            if term in suggested_lower:
                for col in typed_cols:
                    if any(mapping in col.lower() for mapping in mappings):
                        return col
        
        # Fallback: Return first column of expected type
        # Recent fix: Add better logging and avoid inappropriate fallbacks for business terms
        if typed_cols:
            # For business terms like 'sku' or 'revenue', avoid falling back to generic columns
            if suggested_lower in ['sku', 'revenue', 'sales', 'product', 'customer']:
                logger.warning(f"No appropriate match found for business term '{suggested_col}' among available columns: {typed_cols}")
                # Return None to force better handling upstream rather than incorrect mapping
                return None
            logger.debug(f"Using fallback column '{typed_cols[0]}' for suggested '{suggested_col}'")
            return typed_cols[0]
        
        return None
    
    def _create_fallback_result(self, reason: str, df: pd.DataFrame) -> IntentResult:
        """Create a safe fallback result when classification fails."""
        return IntentResult(
            action_type=ActionType.UNKNOWN,
            confidence=0.3,
            confidence_level=ConfidenceLevel.LOW,
            data_complexity=self._assess_data_complexity(df),
            entities=EntityExtraction([], [], [], [], [], [], []),
            reasoning=f"Fallback result: {reason}",
            suggested_columns=self._select_relevant_columns_fallback("", df),
            analysis_plan=None,
            llm_enhanced=False,
            fallback_used=True
        )
    
    def extract_formula_requirements(self, question: str, df: pd.DataFrame, 
                                   context: Optional[str] = None) -> Dict[str, Any]:
        """
        Extract structured formula requirements from natural language queries.
        
        Args:
            question: User's natural language query
            df: DataFrame containing the data for column validation
            context: Optional conversation context
            
        Returns:
            Dictionary containing structured formula requirements:
            {
                "target_columns": List[str],     # Columns to calculate on
                "aggregation_functions": List[str],  # Functions to apply
                "grouping_requirements": Dict,   # Groupby specifications
                "filter_conditions": List[Dict], # WHERE clauses
                "sort_criteria": Dict,          # ORDER BY specifications
                "additional_requirements": Dict, # Complex logic, formulas, etc.
                "confidence": float,            # Extraction confidence
                "query_complexity": str         # simple/medium/complex
            }
        """
        logger.info(f"Extracting formula requirements from: '{question[:100]}...'")
        
        try:
            # Use LLM extraction if available for better accuracy
            if self.openai_client:
                try:
                    return self._llm_extract_formula_requirements(question, df, context)
                except Exception as e:
                    logger.warning(f"LLM formula extraction failed, using rule-based: {str(e)}")
            
            # Fallback to rule-based extraction
            return self._rule_based_extract_formula_requirements(question, df)
            
        except Exception as e:
            logger.error(f"Formula requirements extraction failed: {str(e)}")
            return self._get_default_formula_requirements()
    
    def _llm_extract_formula_requirements(self, question: str, df: pd.DataFrame, 
                                        context: Optional[str] = None) -> Dict[str, Any]:
        """Extract formula requirements using LLM for enhanced accuracy."""
        
        # Build context about available data
        data_context = self._build_data_context(df)
        
        system_prompt = """You are an expert Excel formula analyst. Extract structured requirements from natural language queries about data analysis.

TASK: Parse the user's query and extract specific formula requirements that can be used to generate Excel formulas.

CRITICAL: Return ONLY valid JSON in the exact format specified below.

EXTRACTION TARGETS:
1. TARGET COLUMNS: Which data columns to calculate on (use actual column names from the data)
2. AGGREGATION FUNCTIONS: What calculations to perform (SUM, AVERAGE, COUNT, MAX, MIN, etc.)
3. GROUPING REQUIREMENTS: How to group data (GROUP BY clauses)
4. FILTER CONDITIONS: What data to include/exclude (WHERE clauses)  
5. SORT CRITERIA: How to order results (ORDER BY clauses)
6. ADDITIONAL REQUIREMENTS: Complex logic, custom formulas, conditional logic

AGGREGATION FUNCTIONS TO DETECT:
- Basic: SUM, AVERAGE, COUNT, MAX, MIN, MEDIAN
- Statistical: STDEV, VAR, MODE, PERCENTILE, QUARTILE
- Conditional: SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS, AVERAGEIFS
- Text: CONCATENATE, LEFT, RIGHT, MID, UPPER, LOWER
- Date: YEAR, MONTH, DAY, WEEKDAY, NOW, TODAY
- Lookup: VLOOKUP, XLOOKUP, INDEX, MATCH
- Logical: IF, AND, OR, NOT, IFS
- Mathematical: ROUND, ABS, POWER, SQRT

GROUPING PATTERNS TO DETECT:
- "by [column]" → GROUP BY column
- "for each [category]" → GROUP BY category  
- "per [dimension]" → GROUP BY dimension
- "broken down by" → GROUP BY

FILTER PATTERNS TO DETECT:
- "where [condition]" → WHERE clause
- "for [category] = [value]" → WHERE category = value
- "if [condition]" → conditional logic
- "greater than", "less than", "equal to" → comparison operators
- "between [x] and [y]" → range conditions
- "contains", "starts with", "ends with" → text matching

SORT PATTERNS TO DETECT:
- "order by", "sort by" → ORDER BY
- "top [N]", "bottom [N]" → LIMIT with sort
- "highest", "lowest" → descending/ascending sort
- "ascending", "descending" → sort direction

Return JSON in this exact format:
{
    "target_columns": ["column1", "column2"],
    "aggregation_functions": ["SUM", "AVERAGE"],
    "grouping_requirements": {
        "group_by_columns": ["category_column"],
        "group_by_type": "simple|complex|nested",
        "having_conditions": []
    },
    "filter_conditions": [
        {
            "column": "column_name",
            "operator": "=|>|<|>=|<=|!=|LIKE|BETWEEN|IN",
            "value": "filter_value",
            "logic": "AND|OR"
        }
    ],
    "sort_criteria": {
        "sort_columns": ["column1"],
        "sort_directions": ["ASC|DESC"],
        "limit": null
    },
    "additional_requirements": {
        "conditional_logic": false,
        "custom_formulas": [],
        "date_functions": false,
        "text_functions": false,
        "lookup_functions": false,
        "complex_calculations": false
    },
    "confidence": 0.0-1.0,
    "query_complexity": "simple|medium|complex"
}"""

        user_prompt = f"""Extract formula requirements from this query:

QUERY: "{question}"

AVAILABLE DATA STRUCTURE:
{self._format_data_context_for_extraction(data_context)}

CONTEXT: {context or "No additional context"}

Analyze the query and extract the specific requirements for generating Excel formulas. Focus on:
1. Which columns are mentioned or implied
2. What calculations are requested  
3. How data should be grouped
4. What filters should be applied
5. How results should be sorted
6. Any additional complex requirements

Return the structured JSON response:"""

        try:
            response = self.openai_client.chat.completions.create(
                model=get_model(),
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                max_tokens=1000,
                temperature=0.1
            )
            
            content = response.choices[0].message.content.strip()
            
            # Clean and parse JSON response
            content = content.replace("```json", "").replace("```", "").strip()
            requirements = json.loads(content)
            
            # Validate and enhance the extracted requirements
            validated_requirements = self._validate_formula_requirements(requirements, df)
            
            logger.info("LLM formula requirements extraction successful")
            return validated_requirements
            
        except Exception as e:
            logger.error(f"LLM formula requirements extraction failed: {str(e)}")
            raise
    
    def _rule_based_extract_formula_requirements(self, question: str, df: pd.DataFrame) -> Dict[str, Any]:
        """Rule-based formula requirements extraction as fallback."""
        
        q_lower = question.lower()
        column_names = df.columns.tolist()
        
        # Extract target columns
        target_columns = []
        for col in column_names:
            if col.lower() in q_lower:
                target_columns.append(col)
        
        # If no specific columns mentioned, try to infer from context
        if not target_columns:
            # Look for common column indicators
            for col in column_names:
                col_lower = col.lower()
                if any(indicator in q_lower for indicator in [
                    col_lower, "sales", "revenue", "amount", "quantity", "price", "value"
                ]):
                    target_columns.append(col)
        
        # Extract aggregation functions
        aggregation_functions = []
        agg_patterns = {
            "sum": ["sum", "total", "add"],
            "average": ["average", "avg", "mean"],
            "count": ["count", "number of", "how many"],
            "max": ["max", "maximum", "highest", "largest"],
            "min": ["min", "minimum", "lowest", "smallest"],
            "median": ["median", "middle"],
            "stdev": ["standard deviation", "stdev", "std"]
        }
        
        for func, patterns in agg_patterns.items():
            if any(pattern in q_lower for pattern in patterns):
                aggregation_functions.append(func.upper())
        
        # Extract grouping requirements
        grouping_requirements = {"group_by_columns": [], "group_by_type": "simple", "having_conditions": []}
        
        # Look for grouping indicators
        group_patterns = [r"by (\w+)", r"for each (\w+)", r"per (\w+)", r"grouped by (\w+)", r"broken down by (\w+)"]
        for pattern in group_patterns:
            matches = re.findall(pattern, q_lower)
            for match in matches:
                # Find matching column
                for col in column_names:
                    if match in col.lower() or col.lower() in match:
                        if col not in grouping_requirements["group_by_columns"]:
                            grouping_requirements["group_by_columns"].append(col)
        
        # Extract filter conditions
        filter_conditions = []
        
        # Look for filter patterns
        filter_patterns = [
            (r"where (\w+) (=|>|<|>=|<=|!=) ([^\s]+)", "explicit_where"),
            (r"(\w+) (greater than|less than|equal to) ([^\s]+)", "comparison"),
            (r"(\w+) between ([^\s]+) and ([^\s]+)", "range"),
            (r"(\w+) contains ([^\s]+)", "text_match")
        ]
        
        for pattern, filter_type in filter_patterns:
            matches = re.findall(pattern, q_lower)
            for match in matches:
                if filter_type == "explicit_where":
                    column, operator, value = match
                elif filter_type == "comparison":
                    column, op_text, value = match
                    operator = {"greater than": ">", "less than": "<", "equal to": "="}.get(op_text, "=")
                elif filter_type == "range":
                    column, value1, value2 = match
                    operator, value = "BETWEEN", f"{value1} AND {value2}"
                elif filter_type == "text_match":
                    column, value = match
                    operator = "LIKE"
                
                # Find matching column
                matching_col = None
                for col in column_names:
                    if column in col.lower() or col.lower() in column:
                        matching_col = col
                        break
                
                if matching_col:
                    filter_conditions.append({
                        "column": matching_col,
                        "operator": operator,
                        "value": value,
                        "logic": "AND"
                    })
        
        # Extract sort criteria
        sort_criteria = {"sort_columns": [], "sort_directions": [], "limit": None}
        
        # Look for sorting indicators
        if any(phrase in q_lower for phrase in ["order by", "sort by", "arrange by"]):
            # Try to find the column to sort by
            sort_patterns = [r"(?:order|sort|arrange) by (\w+)", r"(?:order|sort|arrange) (\w+)"]
            for pattern in sort_patterns:
                matches = re.findall(pattern, q_lower)
                for match in matches:
                    for col in column_names:
                        if match in col.lower() or col.lower() in match:
                            sort_criteria["sort_columns"].append(col)
                            # Determine direction
                            if any(word in q_lower for word in ["desc", "descending", "highest", "largest"]):
                                sort_criteria["sort_directions"].append("DESC")
                            else:
                                sort_criteria["sort_directions"].append("ASC")
        
        # Look for top/bottom N
        top_bottom_matches = re.findall(r'(?:top|bottom)\s+(\d+)', q_lower)
        if top_bottom_matches:
            sort_criteria["limit"] = int(top_bottom_matches[0])
            if "bottom" in q_lower:
                sort_criteria["sort_directions"] = ["ASC"]
            else:
                sort_criteria["sort_directions"] = ["DESC"]
        
        # Detect additional requirements
        additional_requirements = {
            "conditional_logic": any(phrase in q_lower for phrase in ["if", "when", "case", "conditional"]),
            "custom_formulas": [],
            "date_functions": any(phrase in q_lower for phrase in ["date", "year", "month", "day", "time"]),
            "text_functions": any(phrase in q_lower for phrase in ["text", "string", "concatenate", "upper", "lower"]),
            "lookup_functions": any(phrase in q_lower for phrase in ["lookup", "vlookup", "match", "index"]),
            "complex_calculations": any(phrase in q_lower for phrase in ["calculate", "formula", "equation", "complex"])
        }
        
        # Assess query complexity
        complexity_score = 0
        complexity_score += len(target_columns)
        complexity_score += len(aggregation_functions) * 2
        complexity_score += len(grouping_requirements["group_by_columns"]) * 2
        complexity_score += len(filter_conditions) * 3
        complexity_score += 2 if sort_criteria["sort_columns"] else 0
        complexity_score += sum(additional_requirements.values()) * 2
        
        if complexity_score <= 3:
            query_complexity = "simple"
            confidence = 0.8
        elif complexity_score <= 8:
            query_complexity = "medium"
            confidence = 0.7
        else:
            query_complexity = "complex"
            confidence = 0.6
        
        return {
            "target_columns": target_columns or [column_names[0]] if column_names else [],
            "aggregation_functions": aggregation_functions or ["SUM"],
            "grouping_requirements": grouping_requirements,
            "filter_conditions": filter_conditions,
            "sort_criteria": sort_criteria,
            "additional_requirements": additional_requirements,
            "confidence": confidence,
            "query_complexity": query_complexity,
            "extraction_method": "rule_based"
        }
    
    def _validate_formula_requirements(self, requirements: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
        """Validate and enhance extracted formula requirements."""
        
        column_names = df.columns.tolist()
        
        # Validate target columns exist
        valid_target_columns = []
        for col in requirements.get("target_columns", []):
            if col in column_names:
                valid_target_columns.append(col)
            else:
                # Try to find similar column names
                similar_cols = [c for c in column_names if col.lower() in c.lower() or c.lower() in col.lower()]
                if similar_cols:
                    valid_target_columns.extend(similar_cols[:1])  # Take first match
        
        requirements["target_columns"] = valid_target_columns
        
        # Validate grouping columns
        grouping_reqs = requirements.get("grouping_requirements", {})
        valid_group_columns = []
        for col in grouping_reqs.get("group_by_columns", []):
            if col in column_names:
                valid_group_columns.append(col)
            else:
                similar_cols = [c for c in column_names if col.lower() in c.lower() or c.lower() in col.lower()]
                if similar_cols:
                    valid_group_columns.extend(similar_cols[:1])
        
        grouping_reqs["group_by_columns"] = valid_group_columns
        requirements["grouping_requirements"] = grouping_reqs
        
        # Validate filter condition columns
        valid_filter_conditions = []
        for condition in requirements.get("filter_conditions", []):
            if condition.get("column") in column_names:
                valid_filter_conditions.append(condition)
            else:
                # Try to find similar column
                col = condition.get("column", "")
                similar_cols = [c for c in column_names if col.lower() in c.lower() or c.lower() in col.lower()]
                if similar_cols:
                    condition["column"] = similar_cols[0]
                    valid_filter_conditions.append(condition)
        
        requirements["filter_conditions"] = valid_filter_conditions
        
        # Validate sort columns
        sort_criteria = requirements.get("sort_criteria", {})
        valid_sort_columns = []
        for col in sort_criteria.get("sort_columns", []):
            if col in column_names:
                valid_sort_columns.append(col)
            else:
                similar_cols = [c for c in column_names if col.lower() in c.lower() or c.lower() in col.lower()]
                if similar_cols:
                    valid_sort_columns.extend(similar_cols[:1])
        
        sort_criteria["sort_columns"] = valid_sort_columns
        requirements["sort_criteria"] = sort_criteria
        
        # Ensure minimum requirements
        if not requirements.get("target_columns") and column_names:
            requirements["target_columns"] = [column_names[0]]
        
        if not requirements.get("aggregation_functions"):
            requirements["aggregation_functions"] = ["SUM"]
        
        # Add validation metadata
        requirements["validation"] = {
            "columns_validated": True,
            "available_columns": column_names,
            "validation_timestamp": pd.Timestamp.now().isoformat()
        }
        
        return requirements
    
    def _format_data_context_for_extraction(self, data_context: Dict[str, Any]) -> str:
        """Format data context for LLM prompt."""
        
        columns_info = []
        if "columns" in data_context:
            for col_name, col_info in data_context["columns"].items():
                col_type = col_info.get("dtype", "unknown")
                sample_values = col_info.get("sample_values", [])
                sample_text = f" (samples: {', '.join(map(str, sample_values[:3]))})" if sample_values else ""
                columns_info.append(f"- {col_name}: {col_type}{sample_text}")
        
        sample_data = data_context.get("sample_data", "No sample data available")
        if isinstance(sample_data, list) and len(sample_data) > 0:
            sample_text = f"Sample rows: {sample_data[:3]}"
        else:
            sample_text = str(sample_data)
        
        return f"""
Columns:
{chr(10).join(columns_info) if columns_info else 'No column information available'}

{sample_text}
        """.strip()
    
    def _get_default_formula_requirements(self) -> Dict[str, Any]:
        """Return default formula requirements when extraction fails."""
        return {
            "target_columns": [],
            "aggregation_functions": ["SUM"],
            "grouping_requirements": {
                "group_by_columns": [],
                "group_by_type": "simple",
                "having_conditions": []
            },
            "filter_conditions": [],
            "sort_criteria": {
                "sort_columns": [],
                "sort_directions": [],
                "limit": None
            },
            "additional_requirements": {
                "conditional_logic": False,
                "custom_formulas": [],
                "date_functions": False,
                "text_functions": False,
                "lookup_functions": False,
                "complex_calculations": False
            },
            "confidence": 0.3,
            "query_complexity": "unknown",
            "extraction_method": "default_fallback"
        }


# Global classifier instance
_intent_classifier = None

def get_intent_classifier() -> EnhancedIntentClassifier:
    """Get or create the global intent classifier instance."""
    global _intent_classifier
    if _intent_classifier is None:
        _intent_classifier = EnhancedIntentClassifier()
    return _intent_classifier


# Backward compatibility functions
def classify_intent(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    """
    Backward compatible intent classification function.
    
    Returns:
      {
        "action_type": "analyze" | "audit" | "chart" | "unknown",
        "confidence": float,
        "data_complexity": "simple" | "medium" | "large" | "massive"
      }
    """
    try:
        classifier = get_intent_classifier()
        result = classifier.classify_intent(question, df)
        
        # Convert to old format for backward compatibility
        return {
            "action_type": result.action_type.value,
            "confidence": result.confidence,
            "data_complexity": result.data_complexity.value
        }
    except Exception as e:
        logger.error(f"Intent classification error: {e}")
        return {
            "action_type": "unknown",
            "confidence": 0.3,
            "data_complexity": "medium"
        }


def _find_best_groupby_column(query_term: str, categorical_columns: List[str]) -> str:
    """
    Find the best column for groupby operations, prioritizing business-meaningful columns over IDs.
    
    Args:
        query_term: The term from the query (e.g., "sku" from "by sku")
        categorical_columns: List of available categorical column names
    
    Returns:
        Best matching column name
    """
    query_term = query_term.lower().strip()
    
    # First, look for exact or partial matches with the query term
    for col in categorical_columns:
        col_lower = col.lower()
        if query_term in col_lower or col_lower in query_term:
            return col
    
    # Second, prioritize business-meaningful columns with preference for specific descriptive terms
    # Recent fix: Prioritize specific terms (sku_code, product_code) over generic terms (id)
    priority_keywords = ['sku_code', 'product_code', 'item_code', 'sku', 'product', 'item', 'category', 'type', 'name', 'code']
    for keyword in priority_keywords:
        for col in categorical_columns:
            if keyword in col.lower():
                return col
    
    # Third, avoid ID-like columns unless no alternative
    id_keywords = ['id', 'uuid', 'guid', 'key']
    non_id_columns = []
    for col in categorical_columns:
        col_lower = col.lower()
        if not any(id_word in col_lower for id_word in id_keywords):
            non_id_columns.append(col)
    
    if non_id_columns:
        return non_id_columns[0]
    
    # Last resort: return first available column
    return categorical_columns[0] if categorical_columns else None


def select_relevant_columns(question: str, header_analysis: Dict[str, Any]) -> Dict[str, Any]:
    """
    Backward compatible column selection function.
    Enhanced with LLM capabilities while maintaining the same interface.
    """
    try:
        # Create a mock DataFrame from header analysis for column selection
        mock_data = {}
        for col_name, col_info in header_analysis.items():
            if isinstance(col_info, dict):
                if col_info.get("is_numeric"):
                    mock_data[col_name] = [1, 2, 3]  # Mock numeric data
                else:
                    mock_data[col_name] = ["A", "B", "C"]  # Mock categorical data
        
        if not mock_data:
            # Fallback if header analysis is malformed
            return {
                "numeric_cols": [],
                "categorical_cols": [],
                "reasoning": "No valid columns found in header analysis"
            }
        
        df = pd.DataFrame(mock_data)
        
        # Use enhanced classifier for column selection
        classifier = get_intent_classifier()
        result = classifier.classify_intent(question, df)
        
        # Enhanced aggregation pattern detection (backward compatibility)
        aggregation_pattern = None
        groupby_pattern = None
        
        # Check if we have aggregations detected
        if result.entities.aggregations:
            aggregation_pattern = result.entities.aggregations[0].upper()  # Ensure uppercase consistency
            logger.debug(f"Found aggregation pattern: {aggregation_pattern}")
            
            # Method 1: Explicit " by " pattern
            if " by " in question.lower():
                logger.debug("Using Method 1: Explicit 'by' pattern")
                parts = question.lower().split(" by ")
                if len(parts) == 2 and result.suggested_columns["categorical"]:
                    # Recent fix: Prioritize business-meaningful columns over ID columns
                    groupby_col = _find_best_groupby_column(parts[1].strip(), result.suggested_columns["categorical"])
                    logger.info(f"Smart column selection: '{parts[1].strip()}' -> '{groupby_col}' from {result.suggested_columns['categorical']}")
                    groupby_pattern = {
                        "metric_column": result.suggested_columns["numeric"][0] if result.suggested_columns["numeric"] else None,
                        "groupby_column": groupby_col,
                        "aggregation": aggregation_pattern
                    }
            
            # Method 2: Complex patterns like "average of X for top N Y"
            elif any(op in result.entities.operations for op in ["top", "bottom", "rank"]):
                logger.debug(f"Using Method 2: Complex patterns with operations: {result.entities.operations}")
                # Pattern: "average of sale for top 10 sku" means group by sku, aggregate sales
                q_lower = question.lower()
                
                # Look for "aggregation [of] metric for top/bottom N dimension" pattern
                # Handles both "average of sale for top 10 sku" and "max revenue for bottom 3 customers"
                complex_pattern = re.search(r'(average|sum|total|max|min|count|mean)\s+(?:of\s+)?(\w+)\s+for\s+(top|bottom)\s+\d+\s+(\w+)', q_lower)
                logger.debug(f"Method 2 complex pattern search: pattern_match={complex_pattern is not None}")
                if complex_pattern:
                    metric_term = complex_pattern.group(2)  # e.g., "sale"
                    dimension_term = complex_pattern.group(4)  # e.g., "sku"
                    
                    # Find matching columns based on terms
                    metric_col = None
                    dimension_col = None
                    
                    # Find metric column (numeric) that matches the term
                    for col in result.suggested_columns["numeric"]:
                        if metric_term in col.lower() or any(term in col.lower() for term in ["sales", "sale", "amount", "value", "revenue"]):
                            metric_col = col
                            break
                    
                    # Find dimension column (categorical) that matches the term  
                    for col in result.suggested_columns["categorical"]:
                        if dimension_term in col.lower() or any(term in col.lower() for term in ["sku", "product", "item"]):
                            dimension_col = col
                            break
                    
                    # Use fallbacks if specific matches not found
                    if not metric_col and result.suggested_columns["numeric"]:
                        metric_col = result.suggested_columns["numeric"][0]
                    if not dimension_col and result.suggested_columns["categorical"]:
                        # Recent fix: Use smart column selection instead of just first column
                        dimension_col = _find_best_groupby_column(dimension_term, result.suggested_columns["categorical"])
                    
                    if metric_col and dimension_col:
                        groupby_pattern = {
                            "metric_column": metric_col,
                            "groupby_column": dimension_col,
                            "aggregation": aggregation_pattern  # Already uppercase
                        }
                        logger.info(f"Detected complex groupby pattern: {aggregation_pattern} {metric_col} by {dimension_col}")
            
            # Method 3: Enhanced fallback for complex patterns without exact regex match
            elif (any(op in result.entities.operations for op in ["top", "bottom", "rank"]) or 
                  any(word in question.lower() for word in ["top", "bottom"]) or
                  result.entities.thresholds) and not groupby_pattern:
                # Handle cases like "max revenue for bottom 3 customers" where regex didn't match perfectly
                q_lower = question.lower()
                logger.debug(f"Entering Method 3 flexible pattern matching for: '{q_lower}'")
                
                # More flexible pattern matching for "AGGREGATION METRIC for top/bottom N DIMENSION"
                flexible_pattern = re.search(r'(max|min|average|sum|total|count|mean)\s+(\w+)\s+for\s+(top|bottom)\s+\d+\s+(\w+)', q_lower)
                logger.debug(f"Flexible pattern search: query='{q_lower}', pattern_match={flexible_pattern is not None}")
                if flexible_pattern:
                    agg_func = flexible_pattern.group(1)
                    metric_term = flexible_pattern.group(2) 
                    rank_type = flexible_pattern.group(3)
                    dimension_term = flexible_pattern.group(4)
                    
                    # Find matching columns
                    metric_col = None
                    dimension_col = None
                    
                    # Find metric column
                    for col in result.suggested_columns["numeric"]:
                        if metric_term in col.lower() or any(term in col.lower() for term in [metric_term, "revenue", "sales", "amount", "value"]):
                            metric_col = col
                            break
                    
                    # Find dimension column
                    for col in result.suggested_columns["categorical"]:
                        if dimension_term in col.lower() or any(term in col.lower() for term in [dimension_term, "customer", "client", "product", "sku"]):
                            dimension_col = col
                            break
                    
                    # Use fallbacks
                    if not metric_col and result.suggested_columns["numeric"]:
                        metric_col = result.suggested_columns["numeric"][0]
                    if not dimension_col and result.suggested_columns["categorical"]:
                        # Recent fix: Use smart column selection instead of just first column
                        dimension_col = _find_best_groupby_column(dimension_term, result.suggested_columns["categorical"])
                    
                    if metric_col and dimension_col:
                        groupby_pattern = {
                            "metric_column": metric_col,
                            "groupby_column": dimension_col,
                            "aggregation": aggregation_pattern  # Already uppercase
                        }
                        logger.info(f"Detected flexible groupby pattern: {aggregation_pattern} {metric_col} by {dimension_col}")
                
                # Fallback: Look for specific entity words
                elif not groupby_pattern:
                    # General fallback for ranking operations with aggregation
                    dimension_col = None
                    metric_col = None
                    
                    # Try to match dimension terms with columns
                    for dim in result.entities.dimensions:
                        for col in result.suggested_columns["categorical"]:
                            if dim in col.lower():
                                dimension_col = col
                                break
                        if dimension_col:
                            break
                    
                    # Try to match metric terms with columns  
                    for metric in result.entities.metrics:
                        for col in result.suggested_columns["numeric"]:
                            if metric in col.lower():
                                metric_col = col
                                break
                        if metric_col:
                            break
                    
                    # Use first available if no specific match
                    if not metric_col and result.suggested_columns["numeric"]:
                        metric_col = result.suggested_columns["numeric"][0]
                    if not dimension_col and result.suggested_columns["categorical"]:
                        dimension_col = result.suggested_columns["categorical"][0]
                    
                    if metric_col and dimension_col:
                        groupby_pattern = {
                            "metric_column": metric_col,
                            "groupby_column": dimension_col,
                            "aggregation": aggregation_pattern
                        }
                        logger.info(f"Detected entity-based groupby pattern: {aggregation_pattern} {metric_col} by {dimension_col}")
            
            # Method 4: Simple aggregation with detected dimensions
            elif result.suggested_columns["categorical"] and not groupby_pattern:
                logger.debug("Using Method 4: Simple aggregation fallback")
                # If we have aggregation and categorical columns but no explicit groupby, create a basic pattern
                groupby_pattern = {
                    "metric_column": result.suggested_columns["numeric"][0] if result.suggested_columns["numeric"] else None,
                    "groupby_column": result.suggested_columns["categorical"][0],
                    "aggregation": aggregation_pattern  # Already uppercase
                }
        
        return {
            "numeric_cols": result.suggested_columns["numeric"],
            "categorical_cols": result.suggested_columns["categorical"],
            "reasoning": result.reasoning,
            "aggregation_pattern": aggregation_pattern,
            "groupby_pattern": groupby_pattern
        }
        
    except Exception as e:
        logger.error(f"Error in enhanced column selection: {e}")
        # Fallback to original simple logic
        all_numeric = [k for k, v in header_analysis.items() if isinstance(v, dict) and v.get("is_numeric")]
        all_categorical = [k for k, v in header_analysis.items() if isinstance(v, dict) and (v.get("is_categorical") or v.get("dtype") == "object")]
        
        return {
            "numeric_cols": all_numeric[:3],
            "categorical_cols": all_categorical[:2],
            "reasoning": "Fallback: used simple column selection due to error"
        }


# New enhanced functions for advanced usage
def classify_intent_enhanced(question: str, df: pd.DataFrame, context: Optional[str] = None) -> IntentResult:
    """
    Enhanced intent classification with full LLM capabilities.
    
    Args:
        question: User's natural language query
        df: DataFrame containing the data
        context: Optional conversation context
        
    Returns:
        IntentResult with comprehensive classification details
    """
    classifier = get_intent_classifier()
    return classifier.classify_intent(question, df, context)


def extract_query_entities(question: str, df: pd.DataFrame) -> EntityExtraction:
    """
    Extract entities from user query using LLM or pattern matching.
    
    Args:
        question: User's natural language query
        df: DataFrame for context
        
    Returns:
        EntityExtraction with identified entities
    """
    classifier = get_intent_classifier()
    result = classifier.classify_intent(question, df)
    return result.entities


def get_analysis_recommendations(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    """
    Get comprehensive analysis recommendations based on intent classification.
    
    Args:
        question: User's natural language query
        df: DataFrame containing the data
        
    Returns:
        Dictionary with analysis recommendations
    """
    classifier = get_intent_classifier()
    result = classifier.classify_intent(question, df)
    
    return {
        "primary_action": result.action_type.value,
        "confidence": result.confidence,
        "confidence_level": result.confidence_level.value,
        "data_complexity": result.data_complexity.value,
        "recommended_columns": result.suggested_columns,
        "extracted_entities": {
            "metrics": result.entities.metrics,
            "dimensions": result.entities.dimensions,
            "aggregations": result.entities.aggregations,
            "time_references": result.entities.time_references,
            "comparisons": result.entities.comparisons,
            "thresholds": result.entities.thresholds,
            "operations": result.entities.operations
        },
        "analysis_plan": result.analysis_plan,
        "reasoning": result.reasoning,
        "llm_enhanced": result.llm_enhanced,
        "fallback_used": result.fallback_used
    }


def get_python_execution_plan(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    """
    Get Python execution plan for complex queries requiring preprocessing.
    
    Args:
        question: User's natural language query
        df: DataFrame containing the data
        
    Returns:
        Dictionary with Python execution plan
    """
    classifier = get_intent_classifier()
    result = classifier.classify_intent(question, df)
    return get_python_execution_plan_from_intent(result)


def get_python_execution_plan_from_intent(intent_result: 'IntentResult') -> Dict[str, Any]:
    """
    Get Python execution plan from already classified intent result.
    
    Recent fix: Avoids redundant intent classification by reusing existing result.
    
    Args:
        intent_result: Already classified IntentResult
        
    Returns:
        Dictionary with Python execution plan
    """
    # Check if Python preprocessing is needed
    analysis_plan = intent_result.analysis_plan or {}
    print(analysis_plan)
    needs_python = analysis_plan.get("python_preprocessing", False)
    
    if not needs_python:
        return {
            "requires_python_preprocessing": False,
            "execution_plan": None,
            "reasoning": "Query can be handled directly with Excel functions"
        }
    
    # Extract execution details
    execution_plan = {
        "requires_python_preprocessing": True,
        "steps": analysis_plan.get("execution_steps", []),
        "ranking_criteria": analysis_plan.get("ranking_criteria"),
        "filter_operation": analysis_plan.get("filter_operation"),
        "filter_value": analysis_plan.get("filter_value"),
        "final_aggregation": analysis_plan.get("final_aggregation"),
        "aggregation_column": analysis_plan.get("aggregation_column"),
        "groupby_column": analysis_plan.get("groupby_column"),
        "suggested_columns": intent_result.suggested_columns,
        "python_code_template": _generate_python_code_template(analysis_plan, intent_result.suggested_columns)
    }
    
    return {
        "requires_python_preprocessing": True,
        "execution_plan": execution_plan,
        "reasoning": f"Complex query requires Python preprocessing: {intent_result.reasoning}"
    }


def _generate_python_code_template(analysis_plan: Dict[str, Any], suggested_columns: Dict[str, List[str]]) -> str:
    """Generate Python code template for the execution plan."""
    
    ranking_col = analysis_plan.get("ranking_criteria", "sales_column")
    filter_op = analysis_plan.get("filter_operation", "top_n")
    filter_val = analysis_plan.get("filter_value", 5)
    agg_col = analysis_plan.get("aggregation_column", "revenue_column")
    group_col = analysis_plan.get("groupby_column", "sku_column")
    final_agg = analysis_plan.get("final_aggregation", "average")
    
    # Map aggregation functions to pandas methods
    agg_mapping = {
        "average": "mean()",
        "sum": "sum()",
        "max": "max()",
        "min": "min()",
        "count": "count()"
    }
    pandas_agg = agg_mapping.get(final_agg, "mean()")
    
    if filter_op == "top_n":
        filter_code = f"df.nlargest({filter_val}, '{ranking_col}')"
    elif filter_op == "bottom_n":
        filter_code = f"df.nsmallest({filter_val}, '{ranking_col}')"
    elif filter_op == "threshold":
        filter_code = f"df[df['{ranking_col}'] > {filter_val}]"
    else:
        filter_code = "df"
    
    template = f"""
# Step 1: Filter data based on ranking criteria
filtered_df = {filter_code}

# Step 2: Apply final aggregation
result = filtered_df.groupby('{group_col}')['{agg_col}'].{pandas_agg}

# Step 3: Prepare for Excel output
result_df = result.reset_index()
result_df.columns = ['{group_col}', '{final_agg}_{agg_col}']
"""
    
    return template.strip()