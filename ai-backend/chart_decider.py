# aibackend/chart_decider.py

from __future__ import annotations

import os
import json
import logging
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    from openai import OpenAI  # v1 client
except Exception:
    OpenAI = None  # type: ignore

from charting import suggest_chart

logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


def _get_openai_client():
    try:
        from llm_config import get_openai_client
        return get_openai_client()
    except Exception as e:
        logger.error(f"OpenAI client init error: {e}")
        return None


def _df_preview(df: pd.DataFrame, max_rows: int = 8) -> str:
    try:
        return df.head(max_rows).to_csv(index=False)
    except Exception:
        return ""


def _validate_llm_choice(headers: List[str], payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not isinstance(payload, dict):
        return None
    
    chart_type = payload.get("chartType")
    chart_title = payload.get("chartTitle") or "AI-Suggested Chart"
    category = payload.get("categoryColumn")
    values = payload.get("valueColumns")

    logger.info(f"Validating LLM chart choice: {payload}")
    logger.info(f"Headers: {headers}")
    # Validate category column
    if category is not None:
        if isinstance(category, list):
            # Filter to only valid headers
            valid_cats = [c for c in category if c in headers]
            if not valid_cats:
                logger.info(f"LLM returned unknown categoryColumn(s): {category}")
                category = None
            else:
                category = valid_cats
        elif isinstance(category, str):
            if category not in headers:
                logger.info(f"LLM returned unknown categoryColumn: {category}")
                category = None

    # Handle the new valueColumns format (explicit column-aggregation pairs)
    validated_series = {}
    
    if isinstance(values, dict):
        valid_aggs = ["SUM", "COUNT", "AVERAGE", "MIN", "MAX", "STDEV", "VAR"]
        
        for series_name, series_config in values.items():
            if isinstance(series_config, dict):
                # New format: {"series_name": {"column": "col", "aggregation": "AGG"}}
                column = series_config.get("column")
                aggregation = series_config.get("aggregation", "SUM")
                
                if column and column in headers:
                    # Validate aggregation, default to SUM if invalid
                    validated_agg = aggregation if aggregation in valid_aggs else "SUM"
                    validated_series[series_name] = {
                        "column": column,
                        "aggregation": validated_agg
                    }
            elif isinstance(series_config, str):
                # Legacy simple format: {"column": "aggregation"}
                if series_name in headers:
                    validated_agg = series_config if series_config in valid_aggs else "SUM"
                    validated_series[series_name] = {
                        "column": series_name,
                        "aggregation": validated_agg
                    }
    
    elif isinstance(values, list):
        # Legacy list format: handle gracefully
        for i, column in enumerate(values):
            if column in headers:
                validated_series[f"{column}_sum"] = {
                    "column": column,
                    "aggregation": "SUM"
                }
    
    if not isinstance(chart_type, str) or not chart_type:
        return None

    inferences = payload.get("inferences", "")

    return {
        "chartType": chart_type,
        "categoryColumn": category,
        "valueColumns": validated_series,
        "chartTitle": chart_title,
        "inferences": inferences,
    }



def choose_chart_via_llm(df: pd.DataFrame, question: str) -> Optional[Dict[str, Any]]:
    client = _get_openai_client()
    if client is None or df is None or df.empty:
        return None

    headers = [str(c) for c in df.columns]
    preview = _df_preview(df, max_rows=8)

    system = (
    "You are a precise charting assistant for Excel. Before answering, work through this step-by-step:\n"
    "- Choose the most appropriate chart type for the user's question and provided tabular headers/preview.\n"
    "- Available aggregation functions: SUM, COUNT, AVERAGE, MIN, MAX, STDEV, VAR.\n"
    "- Pie requires one low-cardinality category and one numeric value field.\n"
    "- Line requires a datetime column plus a numeric value (time series).\n"
    "- ColumnClustered for categorical comparisons; BarClustered when horizontal suits better; XYScatter for paired numerics.\n"
    "- ALWAYS select an appropriate categoryColumn (e.g. Month, Region, Type) when plotting trends or grouping.\n"
    "- CRITICAL TIME RULE: If the user asks for a monthly trend/month-on-month analysis AND the data has BOTH a Year column AND a Month column, "
    "you MUST return categoryColumn as an ARRAY containing BOTH columns, e.g. [\"Year\", \"Month\"]. "
    "This ensures data is grouped per month PER year, not collapsed across all years. "
    "If there is only a single date/month column, return it as a string.\n"
    "- Return STRICT JSON: {chartType, categoryColumn, valueColumns, chartTitle, inferences}.\n"
    "- inferences MUST be a 1-2 sentence analytical forecast or summary of what this chart is designed to reveal.\n"
    "- valueColumns should be an object with unique keys mapping to column-aggregation pairs.\n"
    "- Each key represents a series name, and the value contains the source column and aggregation.\n"
    "- chartType ∈ {Line, ColumnClustered, BarClustered, XYScatter, Pie}."
    )

    user = (
        f"Headers: {headers}\n\n"
        f"Preview (CSV):\n{preview}\n\n"
        f"User question: {question}\n\n"
        "Respond ONLY with JSON, for example:\n"
        "{\n"
        '  "chartType": "ColumnClustered",\n'
        '  "categoryColumn": "ChannelID",\n'
        '  "valueColumns": {\n'
        '    "quantity_sum": {"column": "quantity", "aggregation": "SUM"},\n'
        '    "revenue_average": {"column": "revenue", "aggregation": "AVERAGE"}\n'
        '  },\n'
        '  "chartTitle": "Quantity Sum and Revenue Average by Channel ID",\n'
        '  "inferences": "This chart will highlight the total product movement alongside the average revenue yielded per distinct Channel ID, enabling quick detection of our most lucrative sales channels."\n'
        "}"
    )

    logger.info(f"LLM chart decision prompt (system): {system}")
    logger.info(f"LLM chart decision prompt (user): {user}")
    try:
        from llm_config import get_model
        resp = client.chat.completions.create(
            model=get_model(),
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=1,
            response_format={"type": "json_object"},
        )
        raw = (resp.choices[0].message.content or "").strip()
    except Exception as e:
        logger.error(f"LLM chart decision error: {e}")
        return None

    try:
        parsed = json.loads(raw)
    except Exception as e:
        logger.error(f"LLM chart decision JSON parse error: {e}; content={raw[:200]}")
        return None

    validated = _validate_llm_choice(headers, parsed)
    if not validated:
        logger.info(f"LLM chart decision invalid/incomplete: {parsed}")
        return None

    logger.info(f"LLM chart decision: {validated}")
    return validated


def decide_chart(df: pd.DataFrame, question: Optional[str], use_llm: bool) -> Tuple[str, str, Optional[str], Dict[str, Dict[str, str]], Optional[List[str]], str]:
    """
    Return: chartType, chartTitle, categoryColumn, valueColumns, aggregations, inferences
    """
    if use_llm and question:
        logger.info(f"decide_chart: Using LLM for chart suggestion. Question: {question}")
        llm_plan = choose_chart_via_llm(df, question)
        logger.debug(f"decide_chart: LLM plan: {llm_plan}")
        if llm_plan:
            logger.info(f"decide_chart: LLM returned plan: {llm_plan}")
            return (
                llm_plan["chartType"],
                llm_plan["chartTitle"],
                llm_plan.get("categoryColumn"),
                llm_plan.get("valueColumns", {}),
                None,
                llm_plan.get("inferences", "")
            )
    logger.info("decide_chart: Falling back to heuristic chart suggestion.")
    chart_type, chart_title = suggest_chart(df)
    logger.info(f"decide_chart: Heuristic suggestion: chart_type={chart_type}, chart_title={chart_title}")
    return chart_type, chart_title, None, {}, None, ""


# =========================
# Multi-Chart Coordination Extensions
# =========================

class CoordinatedChartManager:
    """
    Extended chart coordination system for multi-table analysis.
    
    Recent fix: Added coordinated multi-chart functionality with:
    - Chart recommendations based on table relationships
    - Coordinated color schemes across multiple charts
    - Dashboard-style chart layouts
    - Interactive chart connections for filtering
    """
    
    def __init__(self):
        """Initialize coordinated chart manager."""
        self.openai_client = _get_openai_client()
        self.color_palettes = self._initialize_color_palettes()
        logger.info("CoordinatedChartManager initialized")
    
    def plan_coordinated_charts(self, sheets: List[Dict[str, Any]], relationships: List[Dict[str, Any]], 
                              question: str) -> Dict[str, Any]:
        """
        Plan coordinated charts for multiple related tables.
        
        Args:
            sheets: List of analysis sheets from multi-table analysis
            relationships: Table relationships from WorksheetManager
            question: Original user question for context
            
        Returns:
            Coordinated chart plan with layouts, colors, and interactions
        """
        logger.info(f"Planning coordinated charts for {len(sheets)} sheets with {len(relationships)} relationships")
        
        try:
            # Step 1: Analyze each sheet's chart potential
            individual_charts = self._analyze_individual_chart_potential(sheets, question)
            
            # Step 2: Plan relationship-based charts
            relationship_charts = self._plan_relationship_charts(sheets, relationships, question)
            
            # Step 3: Design coordinated color scheme
            color_scheme = self._design_coordinated_colors(individual_charts, relationship_charts)
            
            # Step 4: Plan dashboard layout
            dashboard_layout = self.create_dashboard_layout(individual_charts + relationship_charts, "executive")
            
            # Step 5: Define interactive connections
            interactive_connections = self.define_interactive_connections(individual_charts + relationship_charts, relationships)
            
            coordinated_plan = {
                "individual_charts": individual_charts,
                "relationship_charts": relationship_charts,
                "color_scheme": color_scheme,
                "dashboard_layout": dashboard_layout,
                "interactive_connections": interactive_connections,
                "metadata": {
                    "total_charts": len(individual_charts) + len(relationship_charts),
                    "layout_style": dashboard_layout.get("style", "grid"),
                    "color_palette": color_scheme.get("primary_palette", "business"),
                    "interactivity_level": len(interactive_connections)
                }
            }
            
            logger.info(f"Coordinated chart plan created with {coordinated_plan['metadata']['total_charts']} charts")
            return coordinated_plan
            
        except Exception as e:
            logger.error(f"Failed to plan coordinated charts: {str(e)}")
            return self._create_fallback_chart_plan(sheets, question)
    
    def generate_chart_recommendations_by_relationship(self, source_sheet: Dict[str, Any], 
                                                     target_sheet: Dict[str, Any], 
                                                     relationship: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Generate chart recommendations based on table relationships.
        
        Args:
            source_sheet: Source table data
            target_sheet: Target table data
            relationship: Relationship definition
            
        Returns:
            List of chart recommendations for the relationship
        """
        relationship_type = relationship.get("type")
        recommendations = []
        
        if relationship_type == "summary_to_detail":
            # Summary-detail relationships benefit from comparison charts
            recommendations.extend(self._create_summary_detail_charts(source_sheet, target_sheet, relationship))
            
        elif relationship_type == "cross_analysis":
            # Cross-analysis benefits from side-by-side comparisons
            recommendations.extend(self._create_cross_analysis_charts(source_sheet, target_sheet, relationship))
            
        elif relationship_type == "drill_down":
            # Drill-down benefits from hierarchical visualizations
            recommendations.extend(self._create_drill_down_charts(source_sheet, target_sheet, relationship))
        
        logger.debug(f"Generated {len(recommendations)} chart recommendations for {relationship_type} relationship")
        return recommendations
    
    def apply_coordinated_color_scheme(self, charts: List[Dict[str, Any]], 
                                     color_scheme: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Apply coordinated color schemes across multiple charts.
        
        Args:
            charts: List of chart configurations
            color_scheme: Coordinated color scheme definition
            
        Returns:
            Charts with applied color schemes
        """
        logger.debug(f"Applying coordinated colors to {len(charts)} charts")
        
        primary_colors = color_scheme.get("primary_colors", ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"])
        secondary_colors = color_scheme.get("secondary_colors", ["#aec7e8", "#ffbb78", "#98df8a", "#ff9896"])
        accent_color = color_scheme.get("accent_color", "#9467bd")
        
        for i, chart in enumerate(charts):
            # Apply color scheme based on chart type and position
            if chart.get("chart_type") == "summary":
                # Summary charts get primary colors
                chart["color_scheme"] = {
                    "primary": primary_colors[i % len(primary_colors)],
                    "secondary": secondary_colors[i % len(secondary_colors)],
                    "accent": accent_color
                }
            elif chart.get("chart_type") == "comparison":
                # Comparison charts get contrasting colors
                chart["color_scheme"] = {
                    "primary": primary_colors[0],
                    "secondary": primary_colors[1],
                    "accent": accent_color
                }
            else:
                # Standard charts get sequential colors
                chart["color_scheme"] = {
                    "primary": primary_colors[i % len(primary_colors)],
                    "accent": accent_color
                }
            
            # Add Excel-specific color formatting
            chart["excel_colors"] = self._convert_to_excel_colors(chart["color_scheme"])
        
        logger.info(f"Applied coordinated color scheme to {len(charts)} charts")
        return charts
    
    def create_dashboard_layout(self, charts: List[Dict[str, Any]], 
                              layout_style: str = "executive") -> Dict[str, Any]:
        """
        Create dashboard-style chart layouts.
        
        Args:
            charts: List of chart configurations
            layout_style: Layout style ("executive", "analytical", "operational")
            
        Returns:
            Dashboard layout specification
        """
        logger.info(f"Creating {layout_style} dashboard layout for {len(charts)} charts")
        
        if layout_style == "executive":
            layout = self._create_executive_layout(charts)
        elif layout_style == "analytical":
            layout = self._create_analytical_layout(charts)
        elif layout_style == "operational":
            layout = self._create_operational_layout(charts)
        else:
            layout = self._create_grid_layout(charts)
        
        # Add responsive sizing
        layout["responsive_sizing"] = self._calculate_responsive_sizing(charts, layout)
        
        # Add navigation elements
        layout["navigation"] = self._create_dashboard_navigation(charts)
        
        logger.debug(f"Dashboard layout created with {len(layout.get('panels', []))} panels")
        return layout
    
    def define_interactive_connections(self, charts: List[Dict[str, Any]], 
                                     relationships: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Define interactive chart connections for filtering and drill-down.
        
        Args:
            charts: List of chart configurations
            relationships: Table relationships
            
        Returns:
            List of interactive connection definitions
        """
        logger.debug(f"Defining interactive connections for {len(charts)} charts")
        connections = []
        
        # Create connections based on shared dimensions
        for i, source_chart in enumerate(charts):
            for j, target_chart in enumerate(charts):
                if i != j:
                    connection = self._identify_chart_connection(source_chart, target_chart, relationships)
                    if connection:
                        connections.append(connection)
        
        # Add master-detail connections
        master_detail_connections = self._create_master_detail_connections(charts, relationships)
        connections.extend(master_detail_connections)
        
        # Add filter propagation rules
        for connection in connections:
            connection["filter_rules"] = self._define_filter_propagation_rules(connection)
            connection["excel_implementation"] = self._create_excel_interactivity_specs(connection)
        
        logger.info(f"Defined {len(connections)} interactive connections")
        return connections
    
    # =========================
    # Individual Chart Analysis
    # =========================
    
    def _analyze_individual_chart_potential(self, sheets: List[Dict[str, Any]], question: str) -> List[Dict[str, Any]]:
        """Analyze each sheet for individual chart potential."""
        charts = []
        
        for sheet in sheets:
            if sheet.get("name") == "Dashboard":
                continue  # Skip dashboard sheet
                
            sheet_name = sheet.get("name", "Unknown")
            formulas = sheet.get("formulas", [])
            
            # Analyze sheet content for chart recommendations
            chart_potential = self._assess_sheet_chart_potential(sheet, question)
            
            if chart_potential:
                chart_config = {
                    "sheet_name": sheet_name,
                    "chart_type": chart_potential.get("recommended_type", "ColumnClustered"),
                    "chart_title": f"{sheet_name} Analysis",
                    "data_range": chart_potential.get("data_range", "A:D"),
                    "categories": chart_potential.get("categories", []),
                    "values": chart_potential.get("values", []),
                    "chart_position": chart_potential.get("position", {"row": 1, "col": 5}),
                    "priority": sheet.get("priority", "medium"),
                    "metrics_count": len(formulas)
                }
                charts.append(chart_config)
        
        return charts
    
    def _assess_sheet_chart_potential(self, sheet: Dict[str, Any], question: str) -> Optional[Dict[str, Any]]:
        """Assess a single sheet's potential for charting."""
        formulas = sheet.get("formulas", [])
        
        if len(formulas) == 0:
            return None
        
        # Determine chart type based on formula types and count
        if len(formulas) == 1:
            chart_type = "Pie"  # Single metric
        elif len(formulas) <= 3:
            chart_type = "ColumnClustered"  # Few metrics
        else:
            chart_type = "Line"  # Many metrics, trend-like
        
        # Extract categories and values from formulas
        categories = ["Category"]  # Default category
        values = [f"Value_{i+1}" for i in range(len(formulas))]
        
        return {
            "recommended_type": chart_type,
            "data_range": f"A1:D{len(formulas) + 5}",
            "categories": categories,
            "values": values,
            "position": {"row": len(formulas) + 10, "col": 1}
        }
    
    # =========================
    # Relationship-Based Charts
    # =========================
    
    def _plan_relationship_charts(self, sheets: List[Dict[str, Any]], 
                                relationships: List[Dict[str, Any]], question: str) -> List[Dict[str, Any]]:
        """Plan charts specifically for table relationships."""
        relationship_charts = []
        
        for relationship in relationships:
            source_name = relationship.get("source")
            target_name = relationship.get("target")
            
            # Find source and target sheets
            source_sheet = next((s for s in sheets if s.get("name") == source_name), None)
            target_sheet = next((s for s in sheets if s.get("name") == target_name), None)
            
            if source_sheet and target_sheet:
                charts = self.generate_chart_recommendations_by_relationship(source_sheet, target_sheet, relationship)
                relationship_charts.extend(charts)
        
        return relationship_charts
    
    def _create_summary_detail_charts(self, source_sheet: Dict[str, Any], 
                                    target_sheet: Dict[str, Any], 
                                    relationship: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create charts for summary-to-detail relationships."""
        return [{
            "chart_type": "comparison_bar",
            "chart_title": f"{source_sheet.get('name')} vs {target_sheet.get('name')} Comparison",
            "source_sheet": source_sheet.get("name"),
            "target_sheet": target_sheet.get("name"),
            "relationship_type": "summary_to_detail",
            "linking_columns": relationship.get("linking_columns", []),
            "chart_position": {"row": 1, "col": 10}
        }]
    
    def _create_cross_analysis_charts(self, source_sheet: Dict[str, Any], 
                                    target_sheet: Dict[str, Any], 
                                    relationship: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create charts for cross-analysis relationships."""
        return [{
            "chart_type": "side_by_side_column",
            "chart_title": f"{source_sheet.get('name')} & {target_sheet.get('name')} Cross-Analysis",
            "source_sheet": source_sheet.get("name"),
            "target_sheet": target_sheet.get("name"),
            "relationship_type": "cross_analysis",
            "linking_columns": relationship.get("linking_columns", []),
            "chart_position": {"row": 10, "col": 10}
        }]
    
    def _create_drill_down_charts(self, source_sheet: Dict[str, Any], 
                                target_sheet: Dict[str, Any], 
                                relationship: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create charts for drill-down relationships."""
        return [{
            "chart_type": "hierarchical_bar",
            "chart_title": f"{source_sheet.get('name')} → {target_sheet.get('name')} Drill-Down",
            "source_sheet": source_sheet.get("name"),
            "target_sheet": target_sheet.get("name"),
            "relationship_type": "drill_down",
            "linking_columns": relationship.get("linking_columns", []),
            "chart_position": {"row": 20, "col": 1}
        }]
    
    # =========================
    # Color Coordination
    # =========================
    
    def _initialize_color_palettes(self) -> Dict[str, Dict[str, List[str]]]:
        """Initialize coordinated color palettes for multi-chart use."""
        return {
            "business": {
                "primary_colors": ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"],
                "secondary_colors": ["#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5", "#c49c94"],
                "accent_color": "#e377c2"
            },
            "corporate": {
                "primary_colors": ["#003f5c", "#2f4b7c", "#665191", "#a05195", "#d45087", "#f95d6a"],
                "secondary_colors": ["#7a9cc6", "#8e7ca8", "#a8678f", "#c85e73", "#e85d75", "#ff7c7c"],
                "accent_color": "#ffa600"
            },
            "professional": {
                "primary_colors": ["#264653", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51", "#3d405b"],
                "secondary_colors": ["#6c8ebf", "#82c09a", "#f4e4bc", "#f7d794", "#f0b27a", "#7e8ba3"],
                "accent_color": "#81b29a"
            }
        }
    
    def _design_coordinated_colors(self, individual_charts: List[Dict[str, Any]], 
                                 relationship_charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Design coordinated color scheme for all charts."""
        total_charts = len(individual_charts) + len(relationship_charts)
        
        # Choose palette based on chart count
        if total_charts <= 3:
            palette_name = "business"
        elif total_charts <= 6:
            palette_name = "corporate"
        else:
            palette_name = "professional"
        
        palette = self.color_palettes[palette_name]
        
        return {
            "primary_palette": palette_name,
            "primary_colors": palette["primary_colors"],
            "secondary_colors": palette["secondary_colors"],
            "accent_color": palette["accent_color"],
            "chart_assignments": self._assign_colors_to_charts(individual_charts + relationship_charts, palette)
        }
    
    def _assign_colors_to_charts(self, charts: List[Dict[str, Any]], palette: Dict[str, List[str]]) -> Dict[str, str]:
        """Assign specific colors to each chart."""
        assignments = {}
        primary_colors = palette["primary_colors"]
        
        for i, chart in enumerate(charts):
            chart_id = f"{chart.get('sheet_name', 'chart')}_{chart.get('chart_type', 'default')}"
            assignments[chart_id] = primary_colors[i % len(primary_colors)]
        
        return assignments
    
    def _convert_to_excel_colors(self, color_scheme: Dict[str, str]) -> Dict[str, str]:
        """Convert hex colors to Excel-compatible color specifications."""
        return {
            "primary_rgb": self._hex_to_rgb(color_scheme.get("primary", "#1f77b4")),
            "secondary_rgb": self._hex_to_rgb(color_scheme.get("secondary", "#aec7e8")),
            "accent_rgb": self._hex_to_rgb(color_scheme.get("accent", "#e377c2"))
        }
    
    def _hex_to_rgb(self, hex_color: str) -> str:
        """Convert hex color to RGB string."""
        hex_color = hex_color.lstrip('#')
        return f"RGB({int(hex_color[0:2], 16)}, {int(hex_color[2:4], 16)}, {int(hex_color[4:6], 16)})"
    
    # =========================
    # Dashboard Layouts
    # =========================
    
    def _create_executive_layout(self, charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create executive-style dashboard layout."""
        return {
            "style": "executive",
            "panels": [
                {"type": "summary", "charts": charts[:2], "position": {"row": 1, "col": 1, "span": "full"}},
                {"type": "details", "charts": charts[2:], "position": {"row": 10, "col": 1, "span": "full"}}
            ],
            "title_position": {"row": 1, "col": 1},
            "navigation_position": {"row": 1, "col": 15}
        }
    
    def _create_analytical_layout(self, charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create analytical-style dashboard layout."""
        return {
            "style": "analytical",
            "panels": [
                {"type": "primary", "charts": charts[:1], "position": {"row": 1, "col": 1, "span": "half"}},
                {"type": "secondary", "charts": charts[1:3], "position": {"row": 1, "col": 8, "span": "half"}},
                {"type": "supporting", "charts": charts[3:], "position": {"row": 15, "col": 1, "span": "full"}}
            ],
            "title_position": {"row": 1, "col": 1},
            "filters_position": {"row": 1, "col": 15}
        }
    
    def _create_operational_layout(self, charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create operational-style dashboard layout."""
        # Grid layout for operational dashboards
        cols_per_row = 2
        panels = []
        
        for i, chart in enumerate(charts):
            row = (i // cols_per_row) * 15 + 1
            col = (i % cols_per_row) * 8 + 1
            
            panels.append({
                "type": "metric",
                "charts": [chart],
                "position": {"row": row, "col": col, "span": "quarter"}
            })
        
        return {
            "style": "operational",
            "panels": panels,
            "title_position": {"row": 1, "col": 1},
            "refresh_position": {"row": 1, "col": 15}
        }
    
    def _create_grid_layout(self, charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create simple grid layout."""
        return {
            "style": "grid",
            "panels": [{"type": "chart", "charts": [chart], "position": {"row": i*10+1, "col": 1}} for i, chart in enumerate(charts)],
            "title_position": {"row": 1, "col": 1}
        }
    
    def _calculate_responsive_sizing(self, charts: List[Dict[str, Any]], layout: Dict[str, Any]) -> Dict[str, Any]:
        """Calculate responsive sizing for charts."""
        return {
            "chart_width": 400,
            "chart_height": 300,
            "panel_padding": 20,
            "responsive_breakpoints": {
                "small": {"max_charts_per_row": 1, "chart_width": 300},
                "medium": {"max_charts_per_row": 2, "chart_width": 350},
                "large": {"max_charts_per_row": 3, "chart_width": 400}
            }
        }
    
    def _create_dashboard_navigation(self, charts: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create navigation elements for dashboard."""
        return {
            "type": "tabs",
            "items": [{"name": chart.get("sheet_name", f"Chart {i+1}"), "target": chart.get("sheet_name")} for i, chart in enumerate(charts)],
            "position": {"row": 1, "col": 15}
        }
    
    # =========================
    # Interactive Connections
    # =========================
    
    def _identify_chart_connection(self, source_chart: Dict[str, Any], 
                                 target_chart: Dict[str, Any], 
                                 relationships: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """Identify potential connection between two charts."""
        source_sheet = source_chart.get("sheet_name")
        target_sheet = target_chart.get("sheet_name")
        
        # Find relationship between sheets
        relationship = next((rel for rel in relationships 
                           if (rel.get("source") == source_sheet and rel.get("target") == target_sheet) or
                              (rel.get("source") == target_sheet and rel.get("target") == source_sheet)), None)
        
        if not relationship:
            return None
        
        return {
            "source_chart": source_chart,
            "target_chart": target_chart,
            "connection_type": "filter_propagation",
            "linking_columns": relationship.get("linking_columns", []),
            "interaction_mode": "click_to_filter"
        }
    
    def _create_master_detail_connections(self, charts: List[Dict[str, Any]], 
                                        relationships: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Create master-detail chart connections."""
        connections = []
        
        # Find summary charts that can serve as masters
        summary_charts = [chart for chart in charts if chart.get("priority") == "high"]
        detail_charts = [chart for chart in charts if chart.get("priority") in ["medium", "low"]]
        
        for summary_chart in summary_charts:
            for detail_chart in detail_charts:
                connection = {
                    "source_chart": summary_chart,
                    "target_chart": detail_chart,
                    "connection_type": "master_detail",
                    "interaction_mode": "selection_drives_detail"
                }
                connections.append(connection)
        
        return connections
    
    def _define_filter_propagation_rules(self, connection: Dict[str, Any]) -> Dict[str, Any]:
        """Define how filters propagate between connected charts."""
        return {
            "propagation_mode": "automatic",
            "filter_fields": connection.get("linking_columns", []),
            "update_mode": "real_time",
            "bidirectional": False
        }
    
    def _create_excel_interactivity_specs(self, connection: Dict[str, Any]) -> Dict[str, Any]:
        """Create Excel-specific interactivity specifications."""
        return {
            "implementation": "pivot_table_slicers",
            "slicer_connections": connection.get("linking_columns", []),
            "chart_links": {
                "source": f"={connection['source_chart']['sheet_name']}!$A$1:$D$10",
                "target": f"={connection['target_chart']['sheet_name']}!$A$1:$F$15"
            },
            "conditional_formatting": True
        }
    
    # =========================
    # Fallback and Utilities
    # =========================
    
    def _create_fallback_chart_plan(self, sheets: List[Dict[str, Any]], question: str) -> Dict[str, Any]:
        """Create simple fallback chart plan when advanced planning fails."""
        logger.warning("Creating fallback chart plan")
        
        charts = []
        for i, sheet in enumerate(sheets):
            if sheet.get("name") != "Dashboard":
                charts.append({
                    "sheet_name": sheet.get("name"),
                    "chart_type": "ColumnClustered",
                    "chart_title": f"{sheet.get('name')} Chart",
                    "chart_position": {"row": i*15+1, "col": 5}
                })
        
        return {
            "individual_charts": charts,
            "relationship_charts": [],
            "color_scheme": self.color_palettes["business"],
            "dashboard_layout": self._create_grid_layout(charts),
            "interactive_connections": [],
            "metadata": {"fallback_used": True, "total_charts": len(charts)}
        }


# =========================
# Module-level functions following existing patterns
# =========================

def create_coordinated_charts(sheets: List[Dict[str, Any]], relationships: List[Dict[str, Any]], 
                            question: str) -> Dict[str, Any]:
    """
    Main entry point for coordinated multi-chart creation - following existing module patterns.
    
    Args:
        sheets: List of analysis sheets from multi-table analysis
        relationships: Table relationships from WorksheetManager
        question: Original user question for context
        
    Returns:
        Complete coordinated chart plan with layouts, colors, and interactions
    
    Example:
        >>> sheets = [{"name": "Summary", "formulas": [...]}, {"name": "Detail", "formulas": [...]}]
        >>> relationships = [{"source": "Summary", "target": "Detail", "type": "summary_to_detail"}]
        >>> chart_plan = create_coordinated_charts(sheets, relationships, "sales analysis")
        >>> print(f"Created {chart_plan['metadata']['total_charts']} coordinated charts")
    """
    try:
        manager = CoordinatedChartManager()
        return manager.plan_coordinated_charts(sheets, relationships, question)
    except Exception as e:
        logger.error(f"Coordinated chart creation failed: {str(e)}")
        # Fallback to simple individual charts
        manager = CoordinatedChartManager()
        return manager._create_fallback_chart_plan(sheets, question)


def get_coordinated_chart_manager() -> CoordinatedChartManager:
    """Get or create CoordinatedChartManager instance - following existing patterns."""
    return CoordinatedChartManager()


# =========================
# Integration with Multi-Table Analysis
# =========================

def integrate_charts_with_worksheet_manager(worksheet_result: Dict[str, Any]) -> Dict[str, Any]:
    """
    Integrate coordinated charts with WorksheetManager multi-table analysis results.
    
    Args:
        worksheet_result: Result from WorksheetManager.create_analysis_suite()
        
    Returns:
        Enhanced result with coordinated charts integrated into each sheet
    """
    logger.info("Integrating coordinated charts with multi-table analysis")
    
    try:
        sheets = worksheet_result.get("sheets", [])
        relationships = worksheet_result.get("table_relationships", [])
        question = worksheet_result.get("metadata", {}).get("question", "")
        
        # Create coordinated chart plan
        chart_plan = create_coordinated_charts(sheets, relationships, question)
        
        # Apply charts to individual sheets
        enhanced_sheets = []
        for sheet in sheets:
            enhanced_sheet = dict(sheet)  # Create copy
            
            # Find charts for this sheet
            sheet_charts = []
            for chart in chart_plan.get("individual_charts", []):
                if chart.get("sheet_name") == sheet.get("name"):
                    sheet_charts.append(chart)
            
            # Add chart specifications to sheet
            if sheet_charts:
                enhanced_sheet["charts"] = sheet_charts
                enhanced_sheet["chart_color_scheme"] = chart_plan.get("color_scheme", {})
                
                # Add chart positioning info
                enhanced_sheet["chart_layout"] = {
                    "position": sheet_charts[0].get("chart_position", {"row": 1, "col": 5}),
                    "size": {"width": 400, "height": 300}
                }
            
            enhanced_sheets.append(enhanced_sheet)
        
        # Create enhanced result
        enhanced_result = dict(worksheet_result)
        enhanced_result["sheets"] = enhanced_sheets
        enhanced_result["coordinated_charts"] = chart_plan
        enhanced_result["metadata"]["charts_integrated"] = True
        enhanced_result["metadata"]["total_charts"] = chart_plan.get("metadata", {}).get("total_charts", 0)
        
        logger.info(f"Successfully integrated {chart_plan.get('metadata', {}).get('total_charts', 0)} charts with multi-table analysis")
        return enhanced_result
        
    except Exception as e:
        logger.error(f"Failed to integrate charts with worksheet manager: {str(e)}")
        return worksheet_result  # Return original result if integration fails


# =========================
# Usage Example and Testing
# =========================

if __name__ == "__main__":
    """
    Demonstration of CoordinatedChartManager functionality with multi-table integration.
    Run this file directly to see the coordinated chart system in action.
    """
    
    # Sample multi-table analysis result (mimicking WorksheetManager output)
    sample_worksheet_result = {
        "sheets": [
            {
                "name": "Dashboard",
                "cells": [{"address": "A1", "value": "Multi-Table Analysis Dashboard"}],
                "formulas": []
            },
            {
                "name": "Summary",
                "cells": [{"address": "A1", "value": "Regional Summary"}],
                "formulas": [
                    {"address": "B3", "formula": "=SUM(Sheet1!D:D)", "description": "Total Sales"},
                    {"address": "B4", "formula": "=AVERAGE(Sheet1!E:E)", "description": "Average Profit"}
                ],
                "priority": "high"
            },
            {
                "name": "Detail",
                "cells": [{"address": "A1", "value": "Product Detail Analysis"}],
                "formulas": [
                    {"address": "B3", "formula": "=SUMIF(Sheet1!A:A,A3,Sheet1!D:D)", "description": "Sales by Product"},
                    {"address": "B4", "formula": "=AVERAGEIF(Sheet1!A:A,A3,Sheet1!E:E)", "description": "Avg Profit by Product"},
                    {"address": "B5", "formula": "=COUNTIF(Sheet1!A:A,A3)", "description": "Product Count"}
                ],
                "priority": "medium"
            }
        ],
        "table_relationships": [
            {
                "source": "Summary",
                "target": "Detail", 
                "type": "summary_to_detail",
                "linking_columns": ["region", "product"],
                "strength": 0.8
            }
        ],
        "metadata": {
            "question": "analyze sales and profit by region and product with summary dashboard",
            "total_tables": 2,
            "multi_table_analysis": True
        }
    }
    
    print("=== CoordinatedChartManager Demo ===\n")
    
    try:
        # Test coordinated chart planning
        chart_manager = CoordinatedChartManager()
        sheets = sample_worksheet_result["sheets"]
        relationships = sample_worksheet_result["table_relationships"]
        question = sample_worksheet_result["metadata"]["question"]
        
        print(f"Input: {len(sheets)} sheets, {len(relationships)} relationships")
        print(f"Question: {question}")
        print("-" * 60)
        
        # Plan coordinated charts
        chart_plan = chart_manager.plan_coordinated_charts(sheets, relationships, question)
        
        print("COORDINATED CHART PLAN:")
        print(f"Individual Charts: {len(chart_plan['individual_charts'])}")
        for chart in chart_plan['individual_charts']:
            print(f"  - {chart['sheet_name']}: {chart['chart_type']} ({chart['chart_title']})")
        
        print(f"Relationship Charts: {len(chart_plan['relationship_charts'])}")
        for chart in chart_plan['relationship_charts']:
            print(f"  - {chart['chart_title']} ({chart['relationship_type']})")
        
        print(f"Color Scheme: {chart_plan['color_scheme']['primary_palette']}")
        print(f"Layout Style: {chart_plan['dashboard_layout']['style']}")
        print(f"Interactive Connections: {len(chart_plan['interactive_connections'])}")
        
        print("\n" + "="*60)
        
        # Test integration with worksheet manager
        print("INTEGRATION WITH WORKSHEET MANAGER:")
        enhanced_result = integrate_charts_with_worksheet_manager(sample_worksheet_result)
        
        charts_integrated = enhanced_result["metadata"].get("charts_integrated", False)
        total_charts = enhanced_result["metadata"].get("total_charts", 0)
        
        print(f"Charts Integrated: {charts_integrated}")
        print(f"Total Charts Created: {total_charts}")
        
        # Show enhanced sheets
        for sheet in enhanced_result["sheets"]:
            if "charts" in sheet:
                chart_count = len(sheet["charts"])
                print(f"Sheet '{sheet['name']}': {chart_count} charts added")
        
        print("\n✅ Demo completed successfully!")
        
    except Exception as e:
        print(f"❌ Demo failed: {str(e)}")
        print("Note: This demo requires a complete environment setup for full functionality.")