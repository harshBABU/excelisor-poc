# ai-backend/analysis_coordinator.py

import os
import re
import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

# Local imports following existing patterns
from intent import classify_intent_enhanced, EnhancedIntentClassifier, IntentResult, ActionType
from worksheet_manager import WorksheetManager, _build_worksheet_payload, _to_col_letter
from llm_config import get_model

logger = logging.getLogger(__name__)


class TableRelationshipType(Enum):
    """Types of relationships between analysis tables."""
    SUMMARY_TO_DETAIL = "summary_to_detail"
    CROSS_ANALYSIS = "cross_analysis"
    TIME_SERIES = "time_series"
    HIERARCHICAL = "hierarchical"
    COMPARATIVE = "comparative"
    DRILL_DOWN = "drill_down"


class AnalysisTableType(Enum):
    """Types of analysis tables that can be generated."""
    SUMMARY = "summary"
    DETAIL = "detail"
    PIVOT = "pivot"
    TREND = "trend"
    COMPARISON = "comparison"
    RANKING = "ranking"
    DISTRIBUTION = "distribution"
    CORRELATION = "correlation"


@dataclass
class TableRelationship:
    """Represents a relationship between two analysis tables."""
    source_table: str
    target_table: str
    relationship_type: TableRelationshipType
    description: str
    linking_columns: List[str]
    dependency_level: int  # 0 = independent, 1 = depends on level 0, etc.


@dataclass
class AnalysisTable:
    """Represents a single analysis table with its metadata."""
    name: str
    table_type: AnalysisTableType
    purpose: str
    priority: str  # "high", "medium", "low"
    metrics: List[Dict[str, Any]]
    groupby_columns: List[str]
    filter_conditions: List[Dict[str, Any]]
    chart_recommendations: List[Dict[str, Any]]
    dependency_level: int
    estimated_complexity: str  # "simple", "medium", "complex"
    execution_notes: str


@dataclass
class MultiTableExecutionPlan:
    """Complete execution plan for multi-table analysis."""
    query: str
    intent_result: IntentResult
    tables: List[AnalysisTable]
    relationships: List[TableRelationship]
    execution_order: List[str]  # Table names in order of creation
    chart_coordination: Dict[str, Any]
    metadata: Dict[str, Any]


class MultiTableAnalyzer:
    """
    Advanced analysis coordinator that breaks complex queries into multiple related tables,
    plans their relationships, and coordinates with WorksheetManager for execution.
    
    Recent fix: Created new MultiTableAnalyzer class with support for:
    - Complex query breakdown into multiple analysis tables
    - Table relationship planning (summary → detail → cross-analysis)
    - Coordination with existing WorksheetManager
    - Appropriate chart generation for each table
    - Execution plan showing table relationships
    """
    
    def __init__(self):
        """Initialize MultiTableAnalyzer with dependencies."""
        self.intent_classifier = EnhancedIntentClassifier()
        self.worksheet_manager = WorksheetManager()
        self.max_tables_per_analysis = 8
        self.max_sheet_name_length = 15
        
        # Initialize OpenAI client following recommended pattern
        try:
            from llm_config import get_openai_client
            self.openai_client = get_openai_client()
            logger.info("MultiTableAnalyzer initialized with LLM-powered capabilities")
        except Exception as e:
            logger.warning(f"OpenAI client initialization failed: {e}")
            self.openai_client = None
    
    # =========================
    # Main Analysis Coordination
    # =========================
    
    def analyze_complex_query(self, query: str, data: Dict[str, Any], context: str = "") -> MultiTableExecutionPlan:
        """
        Main entry point: Analyze complex query and create multi-table execution plan.
        
        Args:
            query: User's complex query requiring multi-table analysis
            data: Data context with pattern_sample and metadata
            context: Additional context for analysis
            
        Returns:
            MultiTableExecutionPlan with tables, relationships, and execution order
        """
        logger.info(f"Starting complex query analysis: {query}")
        
        try:
            # Step 1: Enhanced intent classification
            intent_result = self._classify_query_intent(query, data, context)
            logger.debug(f"Intent classification completed: {intent_result.action_type.value}")
            
            # Step 2: Break down query into analysis components
            analysis_components = self._decompose_query_to_components(query, intent_result, data)
            logger.debug(f"Query decomposed into {len(analysis_components)} components")
            
            # Step 3: Plan analysis tables
            tables = self._plan_analysis_tables(analysis_components, data, intent_result)
            logger.debug(f"Planned {len(tables)} analysis tables")
            
            # Step 4: Determine table relationships
            relationships = self._plan_table_relationships(tables, intent_result)
            logger.debug(f"Identified {len(relationships)} table relationships")
            
            # Step 5: Create execution order
            execution_order = self._determine_execution_order(tables, relationships)
            logger.debug(f"Execution order determined: {execution_order}")
            
            # Step 6: Plan chart coordination
            chart_coordination = self._plan_chart_coordination(tables, relationships)
            logger.debug("Chart coordination planned")
            
            # Step 7: Build execution plan
            execution_plan = MultiTableExecutionPlan(
                query=query,
                intent_result=intent_result,
                tables=tables,
                relationships=relationships,
                execution_order=execution_order,
                chart_coordination=chart_coordination,
                metadata={
                    "created": datetime.now().isoformat(),
                    "total_tables": len(tables),
                    "total_relationships": len(relationships),
                    "complexity_level": intent_result.entities.operations,
                    "multi_dimensional": getattr(intent_result, 'analysis_plan', {}).get('is_multi_dimensional', False),
                    "context": context
                }
            )
            
            logger.info(f"Multi-table execution plan created successfully with {len(tables)} tables")
            return execution_plan
            
        except Exception as e:
            logger.error(f"Failed to create multi-table execution plan: {str(e)}")
            # Fallback to single table analysis
            return self._create_fallback_execution_plan(query, data, context)
    
    def execute_multi_table_analysis(self, execution_plan: MultiTableExecutionPlan, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute the multi-table analysis plan using WorksheetManager.
        
        Args:
            execution_plan: MultiTableExecutionPlan to execute
            data: Data context for analysis
            
        Returns:
            Complete workbook structure with multiple coordinated sheets
        """
        logger.info(f"Executing multi-table analysis with {len(execution_plan.tables)} tables")
        
        try:
            sheets = []
            
            # Create overview/summary sheet first
            overview_sheet = self._create_overview_sheet(execution_plan)
            sheets.append(overview_sheet)
            
            # Execute tables in dependency order
            for table_name in execution_plan.execution_order:
                table = next((t for t in execution_plan.tables if t.name == table_name), None)
                if table:
                    logger.info(f"Creating table: {table.name}")
                    sheet = self._create_table_sheet(table, data, execution_plan)
                    sheets.append(sheet)
            
            # Create relationship visualization sheet if complex enough
            if len(execution_plan.relationships) > 2:
                relationship_sheet = self._create_relationship_visualization_sheet(execution_plan)
                sheets.append(relationship_sheet)
            
            result = {
                "sheets": sheets,
                "execution_plan": {
                    "query": execution_plan.query,
                    "total_tables": len(execution_plan.tables),
                    "analysis_focus": execution_plan.intent_result.reasoning,
                    "table_relationships": [
                        {
                            "from": rel.source_table,
                            "to": rel.target_table,
                            "type": rel.relationship_type.value,
                            "description": rel.description
                        }
                        for rel in execution_plan.relationships
                    ]
                },
                "metadata": {
                    **execution_plan.metadata,
                    "executed": datetime.now().isoformat(),
                    "total_sheets_created": len(sheets),
                    "multi_table_analysis": True
                }
            }
            
            logger.info(f"Multi-table analysis executed successfully with {len(sheets)} sheets")
            return result
            
        except Exception as e:
            logger.error(f"Failed to execute multi-table analysis: {str(e)}")
            # Fallback to WorksheetManager default behavior
            return self.worksheet_manager.create_llm_driven_workbook(data, execution_plan.query)
    
    # =========================
    # Query Analysis & Decomposition
    # =========================
    
    def _classify_query_intent(self, query: str, data: Dict[str, Any], context: str) -> IntentResult:
        """Classify query intent using enhanced intent classifier."""
        # Create a simple DataFrame for intent classification
        pattern_sample = data.get('pattern_sample', [])
        if pattern_sample and len(pattern_sample) > 1:
            headers = pattern_sample[0]
            import pandas as pd
            df = pd.DataFrame(pattern_sample[1:], columns=headers)
        else:
            import pandas as pd
            df = pd.DataFrame()
        
        return self.intent_classifier.classify_intent(query, df, context)
    
    def _decompose_query_to_components(self, query: str, intent_result: IntentResult, data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Break down complex query into analysis components using LLM.
        """
        if not self.openai_client:
            return self._fallback_decompose_query(query, intent_result)
        
        try:
            system_prompt = """You are an expert data analyst that breaks down complex queries into multiple analysis components.

Your task is to analyze a user query and determine what separate analysis tables would be needed to fully answer it.

For each component, identify:
1. The specific analysis focus
2. Required metrics and calculations
3. Grouping dimensions
4. Data filtering needs
5. How it relates to other components

Return a JSON array of analysis components:
[
    {
        "component_name": "descriptive_name",
        "analysis_focus": "what this component analyzes",
        "required_metrics": ["metric1", "metric2"],
        "aggregation_functions": ["sum", "average", "count"],
        "groupby_dimensions": ["dimension1", "dimension2"],
        "filter_conditions": [{"column": "col", "operator": "=", "value": "val"}],
        "table_type": "summary|detail|pivot|trend|comparison|ranking",
        "priority": "high|medium|low",
        "depends_on": ["component_name"] // if depends on other components
    }
]"""
            
            # Build data context
            data_context = self._build_data_context_for_llm(data)
            
            user_prompt = f"""QUERY: {query}

INTENT ANALYSIS:
- Action Type: {intent_result.action_type.value}
- Multi-dimensional: {getattr(intent_result, 'analysis_plan', {}).get('is_multi_dimensional', False)}
- Aggregations: {intent_result.entities.aggregations}
- Dimensions: {intent_result.entities.dimensions}
- Operations: {intent_result.entities.operations}

DATA CONTEXT:
{data_context}

Please break this query down into analysis components that would require separate tables."""
            
            response = self.openai_client.chat.completions.create(
                model=get_model(),
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                max_tokens=2000
            )
            
            content = response.choices[0].message.content.strip()
            logger.debug(f"LLM query decomposition response: {content}")
            
            # Parse JSON response
            components = json.loads(content)
            logger.info(f"Successfully decomposed query into {len(components)} components")
            return components
            
        except Exception as e:
            logger.warning(f"LLM query decomposition failed: {str(e)}, using fallback")
            return self._fallback_decompose_query(query, intent_result)
    
    def _fallback_decompose_query(self, query: str, intent_result: IntentResult) -> List[Dict[str, Any]]:
        """Fallback query decomposition using pattern matching."""
        components = []
        
        # Extract analysis plan if available
        analysis_plan = getattr(intent_result, 'analysis_plan', {})
        is_multi_dim = analysis_plan.get('is_multi_dimensional', False)
        multi_agg = analysis_plan.get('multi_aggregation', False)
        multi_groupby = analysis_plan.get('multi_groupby', False)
        
        if is_multi_dim:
            # Multi-dimensional query - create multiple components
            if multi_agg and multi_groupby:
                # Summary component
                components.append({
                    "component_name": "summary_analysis",
                    "analysis_focus": "Overall summary metrics",
                    "required_metrics": intent_result.entities.metrics[:2],  # First 2 metrics
                    "aggregation_functions": intent_result.entities.aggregations[:2],  # First 2 aggregations
                    "groupby_dimensions": intent_result.entities.dimensions,
                    "filter_conditions": [],
                    "table_type": "summary",
                    "priority": "high",
                    "depends_on": []
                })
                
                # Detail component
                components.append({
                    "component_name": "detailed_breakdown",
                    "analysis_focus": "Detailed breakdown by dimensions",
                    "required_metrics": intent_result.entities.metrics,
                    "aggregation_functions": intent_result.entities.aggregations,
                    "groupby_dimensions": intent_result.entities.dimensions,
                    "filter_conditions": [],
                    "table_type": "detail",
                    "priority": "medium",
                    "depends_on": ["summary_analysis"]
                })
            else:
                # Single component for simpler multi-dimensional queries
                components.append({
                    "component_name": "multi_analysis",
                    "analysis_focus": "Multi-dimensional analysis",
                    "required_metrics": intent_result.entities.metrics,
                    "aggregation_functions": intent_result.entities.aggregations,
                    "groupby_dimensions": intent_result.entities.dimensions,
                    "filter_conditions": [],
                    "table_type": "pivot",
                    "priority": "high",
                    "depends_on": []
                })
        else:
            # Single-dimensional query - one component
            components.append({
                "component_name": "main_analysis",
                "analysis_focus": f"{intent_result.action_type.value} analysis",
                "required_metrics": intent_result.entities.metrics,
                "aggregation_functions": intent_result.entities.aggregations,
                "groupby_dimensions": intent_result.entities.dimensions,
                "filter_conditions": [],
                "table_type": "summary",
                "priority": "high",
                "depends_on": []
            })
        
        return components
    
    # =========================
    # Table Planning & Relationships
    # =========================
    
    def _plan_analysis_tables(self, components: List[Dict[str, Any]], data: Dict[str, Any], intent_result: IntentResult) -> List[AnalysisTable]:
        """Convert analysis components into structured AnalysisTable objects."""
        tables = []
        
        for i, component in enumerate(components):
            # Generate safe sheet name
            base_name = component.get('component_name', f'Table{i+1}')
            sheet_name = self._generate_safe_sheet_name(base_name)
            
            # Create metrics list in WorksheetManager format
            metrics = []
            required_metrics = component.get('required_metrics', [])
            aggregation_functions = component.get('aggregation_functions', [])
            
            for j, metric in enumerate(required_metrics):
                agg_func = aggregation_functions[j] if j < len(aggregation_functions) else 'SUM'
                metrics.append({
                    "name": f"{agg_func.title()} of {metric}",
                    "formula_type": agg_func.upper(),
                    "target_column": metric,
                    "description": f"Calculate {agg_func} of {metric}",
                    "filter_condition": None
                })
            
            # Plan chart recommendations
            chart_recommendations = self._plan_table_charts(component, len(required_metrics))
            
            # Determine dependency level
            dependency_level = 0
            depends_on = component.get('depends_on', [])
            if depends_on:
                dependency_level = max([
                    next((idx for idx, comp in enumerate(components) if comp.get('component_name') in depends_on), 0)
                    for _ in depends_on
                ]) + 1
            
            table = AnalysisTable(
                name=sheet_name,
                table_type=AnalysisTableType(component.get('table_type', 'summary')),
                purpose=component.get('analysis_focus', 'Data analysis'),
                priority=component.get('priority', 'medium'),
                metrics=metrics,
                groupby_columns=component.get('groupby_dimensions', []),
                filter_conditions=component.get('filter_conditions', []),
                chart_recommendations=chart_recommendations,
                dependency_level=dependency_level,
                estimated_complexity=self._estimate_table_complexity(component),
                execution_notes=f"Component {i+1}: {component.get('analysis_focus', 'Analysis')}"
            )
            
            tables.append(table)
        
        return tables
    
    def _plan_table_relationships(self, tables: List[AnalysisTable], intent_result: IntentResult) -> List[TableRelationship]:
        """Determine relationships between analysis tables."""
        relationships = []
        
        # Sort tables by dependency level
        sorted_tables = sorted(tables, key=lambda t: t.dependency_level)
        
        for i, table in enumerate(sorted_tables):
            for j, other_table in enumerate(sorted_tables):
                if i != j and table.dependency_level < other_table.dependency_level:
                    # Find potential relationships
                    relationship = self._identify_table_relationship(table, other_table)
                    if relationship:
                        relationships.append(relationship)
        
        # Add cross-analysis relationships for tables at same level
        same_level_tables = {}
        for table in tables:
            level = table.dependency_level
            if level not in same_level_tables:
                same_level_tables[level] = []
            same_level_tables[level].append(table)
        
        for level, level_tables in same_level_tables.items():
            if len(level_tables) > 1:
                for i, table1 in enumerate(level_tables):
                    for table2 in level_tables[i+1:]:
                        if self._have_common_dimensions(table1, table2):
                            relationships.append(TableRelationship(
                                source_table=table1.name,
                                target_table=table2.name,
                                relationship_type=TableRelationshipType.CROSS_ANALYSIS,
                                description=f"Cross-analysis between {table1.table_type.value} and {table2.table_type.value}",
                                linking_columns=list(set(table1.groupby_columns) & set(table2.groupby_columns)),
                                dependency_level=level
                            ))
        
        return relationships
    
    def _identify_table_relationship(self, source_table: AnalysisTable, target_table: AnalysisTable) -> Optional[TableRelationship]:
        """Identify the relationship type between two tables."""
        # Summary to Detail relationship
        if source_table.table_type == AnalysisTableType.SUMMARY and target_table.table_type == AnalysisTableType.DETAIL:
            return TableRelationship(
                source_table=source_table.name,
                target_table=target_table.name,
                relationship_type=TableRelationshipType.SUMMARY_TO_DETAIL,
                description=f"Summary table feeds into detailed analysis",
                linking_columns=list(set(source_table.groupby_columns) & set(target_table.groupby_columns)),
                dependency_level=target_table.dependency_level
            )
        
        # Drill-down relationship
        if len(source_table.groupby_columns) < len(target_table.groupby_columns):
            return TableRelationship(
                source_table=source_table.name,
                target_table=target_table.name,
                relationship_type=TableRelationshipType.DRILL_DOWN,
                description=f"Drill down from {source_table.table_type.value} to more detailed view",
                linking_columns=source_table.groupby_columns,
                dependency_level=target_table.dependency_level
            )
        
        return None
    
    def _determine_execution_order(self, tables: List[AnalysisTable], relationships: List[TableRelationship]) -> List[str]:
        """Determine the optimal order for creating tables based on dependencies."""
        # Sort by dependency level first, then by priority
        priority_weight = {"high": 3, "medium": 2, "low": 1}
        
        sorted_tables = sorted(tables, key=lambda t: (
            t.dependency_level,
            -priority_weight.get(t.priority, 1),  # Negative for descending order
            t.name
        ))
        
        return [table.name for table in sorted_tables]
    
    # =========================
    # Chart Coordination
    # =========================
    
    def _plan_chart_coordination(self, tables: List[AnalysisTable], relationships: List[TableRelationship]) -> Dict[str, Any]:
        """Plan coordinated charting across multiple tables."""
        chart_plan = {
            "individual_charts": {},
            "relationship_charts": [],
            "dashboard_recommendations": [],
            "chart_coordination_notes": []
        }
        
        # Individual table charts
        for table in tables:
            chart_plan["individual_charts"][table.name] = table.chart_recommendations
        
        # Relationship-based charts
        for relationship in relationships:
            if relationship.relationship_type == TableRelationshipType.SUMMARY_TO_DETAIL:
                chart_plan["relationship_charts"].append({
                    "type": "comparison_chart",
                    "tables": [relationship.source_table, relationship.target_table],
                    "description": f"Compare summary vs detail metrics",
                    "chart_type": "side_by_side_bar"
                })
        
        # Dashboard recommendations
        if len(tables) >= 3:
            chart_plan["dashboard_recommendations"].append({
                "type": "executive_dashboard",
                "primary_table": tables[0].name,
                "supporting_tables": [t.name for t in tables[1:3]],
                "layout": "summary_top_details_bottom"
            })
        
        return chart_plan
    
    def _plan_table_charts(self, component: Dict[str, Any], metric_count: int) -> List[Dict[str, Any]]:
        """Plan appropriate charts for a single table."""
        charts = []
        table_type = component.get('table_type', 'summary')
        groupby_dims = component.get('groupby_dimensions', [])
        
        if table_type == 'summary' and len(groupby_dims) == 1:
            charts.append({
                "type": "column_chart",
                "description": f"Column chart showing metrics by {groupby_dims[0]}",
                "primary": True
            })
        elif table_type == 'comparison' and len(groupby_dims) >= 2:
            charts.append({
                "type": "clustered_column",
                "description": f"Clustered column chart comparing across dimensions",
                "primary": True
            })
        elif table_type == 'trend':
            charts.append({
                "type": "line_chart",
                "description": "Line chart showing trends over time",
                "primary": True
            })
        elif table_type == 'ranking':
            charts.append({
                "type": "horizontal_bar",
                "description": "Horizontal bar chart showing rankings",
                "primary": True
            })
        
        # Add secondary chart if multiple metrics
        if metric_count > 1:
            charts.append({
                "type": "combo_chart",
                "description": "Combination chart for multiple metrics",
                "primary": False
            })
        
        return charts
    
    # =========================
    # Sheet Creation & Execution
    # =========================
    
    def _create_overview_sheet(self, execution_plan: MultiTableExecutionPlan) -> Dict[str, Any]:
        """Create overview sheet explaining the multi-table analysis."""
        cells = []
        row = 1
        
        # Title
        cells.append({"address": f"A{row}", "value": "Multi-Table Analysis Overview"})
        row += 2
        
        # Query
        cells.append({"address": f"A{row}", "value": "Original Query:"})
        cells.append({"address": f"B{row}", "value": execution_plan.query})
        row += 2
        
        # Analysis plan
        cells.append({"address": f"A{row}", "value": "Analysis Approach:"})
        cells.append({"address": f"B{row}", "value": execution_plan.intent_result.reasoning})
        row += 2
        
        # Tables overview
        cells.append({"address": f"A{row}", "value": "Analysis Tables:"})
        row += 1
        
        for i, table in enumerate(execution_plan.tables):
            cells.append({"address": f"B{row}", "value": f"{i+1}. {table.name}"})
            cells.append({"address": f"C{row}", "value": table.purpose})
            cells.append({"address": f"D{row}", "value": f"Type: {table.table_type.value}"})
            row += 1
        
        row += 1
        
        # Relationships
        if execution_plan.relationships:
            cells.append({"address": f"A{row}", "value": "Table Relationships:"})
            row += 1
            
            for rel in execution_plan.relationships:
                cells.append({"address": f"B{row}", "value": f"{rel.source_table} → {rel.target_table}"})
                cells.append({"address": f"C{row}", "value": rel.relationship_type.value})
                cells.append({"address": f"D{row}", "value": rel.description})
                row += 1
        
        return _build_worksheet_payload(
            name="Overview",
            cells=cells,
            notes="Multi-table analysis overview and execution plan"
        )
    
    def _create_table_sheet(self, table: AnalysisTable, data: Dict[str, Any], execution_plan: MultiTableExecutionPlan) -> Dict[str, Any]:
        """Create a sheet for a specific analysis table using WorksheetManager."""
        # Convert AnalysisTable to WorksheetManager sheet plan format
        sheet_plan = {
            "name": table.name,
            "purpose": table.purpose,
            "metrics": table.metrics,
            "priority": table.priority
        }
        
        # Use WorksheetManager to create the actual sheet
        header_map = self.worksheet_manager.map_headers_to_letters(data)
        return self.worksheet_manager._create_sheet_from_llm_plan(sheet_plan, data, header_map)
    
    def _create_relationship_visualization_sheet(self, execution_plan: MultiTableExecutionPlan) -> Dict[str, Any]:
        """Create a sheet visualizing table relationships."""
        cells = []
        row = 1
        
        # Title
        cells.append({"address": f"A{row}", "value": "Table Relationship Map"})
        row += 2
        
        # Create a simple text-based relationship diagram
        cells.append({"address": f"A{row}", "value": "Execution Flow:"})
        row += 1
        
        for i, table_name in enumerate(execution_plan.execution_order):
            indent = "  " * (i % 3)  # Simple indentation
            cells.append({"address": f"A{row}", "value": f"{indent}{i+1}. {table_name}"})
            
            # Find relationships for this table
            for rel in execution_plan.relationships:
                if rel.source_table == table_name:
                    cells.append({"address": f"B{row}", "value": f"→ {rel.target_table} ({rel.relationship_type.value})"})
                    break
            
            row += 1
        
        return _build_worksheet_payload(
            name="Relationships",
            cells=cells,
            notes="Visual representation of table relationships and execution flow"
        )
    
    # =========================
    # Utility Methods
    # =========================
    
    def _generate_safe_sheet_name(self, base_name: str) -> str:
        """Generate Excel-safe sheet name following WorksheetManager patterns."""
        # Remove invalid characters
        safe_name = re.sub(r'[^\w\s-]', '', base_name)
        safe_name = safe_name.replace(' ', '_')
        
        # Truncate to max length
        if len(safe_name) > self.max_sheet_name_length:
            safe_name = safe_name[:self.max_sheet_name_length]
        
        return safe_name or "Analysis"
    
    def _build_data_context_for_llm(self, data: Dict[str, Any]) -> str:
        """Build data context string for LLM analysis."""
        pattern_sample = data.get('pattern_sample', [])
        if not pattern_sample:
            return "No data sample available"
        
        headers = pattern_sample[0] if pattern_sample else []
        sample_rows = pattern_sample[1:6] if len(pattern_sample) > 1 else []  # First 5 rows
        
        context_parts = [
            f"Columns: {', '.join(headers)}",
            f"Sample data ({len(sample_rows)} rows):"
        ]
        
        for row in sample_rows:
            context_parts.append(f"  {row}")
        
        return "\n".join(context_parts)
    
    def _estimate_table_complexity(self, component: Dict[str, Any]) -> str:
        """Estimate complexity of a table component."""
        metrics_count = len(component.get('required_metrics', []))
        dimensions_count = len(component.get('groupby_dimensions', []))
        filters_count = len(component.get('filter_conditions', []))
        
        total_complexity = metrics_count + dimensions_count + filters_count
        
        if total_complexity <= 3:
            return "simple"
        elif total_complexity <= 6:
            return "medium"
        else:
            return "complex"
    
    def _have_common_dimensions(self, table1: AnalysisTable, table2: AnalysisTable) -> bool:
        """Check if two tables have common grouping dimensions."""
        return bool(set(table1.groupby_columns) & set(table2.groupby_columns))
    
    def _create_fallback_execution_plan(self, query: str, data: Dict[str, Any], context: str) -> MultiTableExecutionPlan:
        """Create a simple fallback execution plan when LLM analysis fails."""
        logger.warning("Creating fallback execution plan")
        
        # Simple intent classification
        intent_result = self._classify_query_intent(query, data, context)
        
        # Create single table
        table = AnalysisTable(
            name="Analysis",
            table_type=AnalysisTableType.SUMMARY,
            purpose="Primary analysis table",
            priority="high",
            metrics=[],
            groupby_columns=[],
            filter_conditions=[],
            chart_recommendations=[{"type": "column_chart", "description": "Basic chart", "primary": True}],
            dependency_level=0,
            estimated_complexity="simple",
            execution_notes="Fallback single-table analysis"
        )
        
        return MultiTableExecutionPlan(
            query=query,
            intent_result=intent_result,
            tables=[table],
            relationships=[],
            execution_order=["Analysis"],
            chart_coordination={"individual_charts": {"Analysis": table.chart_recommendations}},
            metadata={
                "created": datetime.now().isoformat(),
                "total_tables": 1,
                "fallback_used": True,
                "context": context
            }
        )


# =========================
# Module-level functions following existing patterns
# =========================

def create_multi_table_analysis(query: str, data: Dict[str, Any], context: str = "") -> Dict[str, Any]:
    """
    Main entry point for multi-table analysis - following existing module patterns.
    
    Args:
        query: Complex user query requiring multi-table analysis
        data: Data context with pattern_sample and metadata
        context: Additional context for analysis
        
    Returns:
        Complete workbook structure with coordinated multi-table analysis
    
    Example:
        >>> data = {
        ...     'pattern_sample': [
        ...         ['product', 'region', 'sales', 'quantity'],
        ...         ['A', 'North', 1000, 50],
        ...         ['B', 'South', 1500, 75]
        ...     ]
        ... }
        >>> query = "sum sales and average quantity by product and region with top 5 products"
        >>> result = create_multi_table_analysis(query, data)
        >>> print(f"Created {len(result['sheets'])} coordinated analysis sheets")
    """
    try:
        analyzer = MultiTableAnalyzer()
        execution_plan = analyzer.analyze_complex_query(query, data, context)
        return analyzer.execute_multi_table_analysis(execution_plan, data)
    except Exception as e:
        logger.error(f"Multi-table analysis failed: {str(e)}")
        # Fallback to standard WorksheetManager
        worksheet_manager = WorksheetManager()
        return worksheet_manager.create_llm_driven_workbook(data, query, context)


def get_multi_table_analyzer() -> MultiTableAnalyzer:
    """Get or create MultiTableAnalyzer instance - following existing patterns."""
    return MultiTableAnalyzer()


# =========================
# Usage Example and Testing
# =========================

if __name__ == "__main__":
    """
    Demonstration of MultiTableAnalyzer functionality.
    Run this file directly to see the multi-table analysis in action.
    """
    
    # Sample data for demonstration
    sample_data = {
        "pattern_sample": [
            ["product_id", "region", "category", "sales", "quantity", "profit", "customer_type"],
            ["P001", "North", "Electronics", 15000, 150, 3000, "Premium"],
            ["P002", "South", "Electronics", 12000, 120, 2400, "Standard"],
            ["P003", "North", "Clothing", 8000, 200, 1600, "Premium"],
            ["P004", "East", "Electronics", 18000, 180, 3600, "Premium"],
            ["P005", "West", "Clothing", 9500, 190, 1900, "Standard"],
            ["P006", "South", "Home", 11000, 110, 2200, "Premium"],
        ],
        "original_size": 1000,
        "headers": ["product_id", "region", "category", "sales", "quantity", "profit", "customer_type"]
    }
    
    # Complex query examples
    test_queries = [
        "sum sales and average profit by region and category with top 5 regions",
        "count products and total quantity by customer type and category",
        "show total sales, average profit, and count of products grouped by region",
        "analyze sales performance across regions with detailed breakdown by product category"
    ]
    
    print("=== MultiTableAnalyzer Demo ===\n")
    
    try:
        analyzer = MultiTableAnalyzer()
        
        for i, query in enumerate(test_queries, 1):
            print(f"Query {i}: {query}")
            print("-" * 60)
            
            # Analyze the query
            execution_plan = analyzer.analyze_complex_query(query, sample_data)
            
            print(f"Tables planned: {len(execution_plan.tables)}")
            for table in execution_plan.tables:
                print(f"  - {table.name}: {table.purpose} ({table.table_type.value})")
            
            print(f"Relationships: {len(execution_plan.relationships)}")
            for rel in execution_plan.relationships:
                print(f"  - {rel.source_table} → {rel.target_table} ({rel.relationship_type.value})")
            
            print(f"Execution order: {' → '.join(execution_plan.execution_order)}")
            print(f"Multi-dimensional: {execution_plan.metadata.get('multi_dimensional', False)}")
            print()
            
    except Exception as e:
        print(f"Demo failed: {str(e)}")
        print("Note: This demo requires OpenAI API key for full LLM functionality.")
