# aibackend/worksheet_manager.py

import os
import re
import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional
from collections import Counter
from intent import select_relevant_columns
from llm_config import get_model

logger = logging.getLogger(__name__)

def _to_col_letter(idx: int) -> str:
    """0-based index to Excel column letters."""
    idx += 1
    letters = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        letters = chr(65 + r) + letters
    return letters

def _build_worksheet_payload(
    name: str,
    cells: List[Dict[str, Any]] = None,
    formulas: List[Dict[str, str]] = None,
    notes: str = ""
) -> Dict[str, Any]:
    """
    Construct a normalized worksheet payload for frontends:
    - cells: [{"address":"A1","value":"Title"}, ...]
    - formulas: [{"address":"F2","formula":"=SUM(A2:A101)"}, ...]
    """
    logger.debug(f"Building worksheet payload: {name}")
    
    cells = cells or []
    formulas = formulas or []
    payload = {"name": name or "AI_Output", "cells": [], "formulas": []}

    if notes:
        payload["cells"].append({"address": "A1", "value": "AI Notes"})
        payload["cells"].append({"address": "A2", "value": notes})

    for c in cells:
        payload["cells"].append({
            "address": c.get("address", "A1"),
            "value": c.get("value", "")
        })

    for f in formulas:
        addr = f.get("address") or f.get("cell") or "A1"
        form = f.get("formula", "")
        if form and not str(form).startswith("="):
            form = "=" + str(form)
        payload["formulas"].append({
            "address": addr,
            "formula": form,
            "description": f.get("description", "")
        })


    return payload

class WorksheetManager:
    """
    LLM-driven worksheet manager that intelligently creates Excel analysis structures
    based on AI understanding of data patterns and user requirements.
    """

    def __init__(self):
        self.max_sheet_name_length = 15
        self.reserved_names = ['Sheet1', 'Sheet2', 'Sheet3']
        
        try:
            from llm_config import get_openai_client
            self.openai_client = get_openai_client()
            logger.info("WorksheetManager initialized with LLM-driven capabilities")
        except Exception as e:
            logger.warning(f"OpenAI client initialization failed: {e}")
            self.openai_client = None

    # =========================
    # LLM-Driven Analysis System
    # =========================

    def llm_decide_metrics_and_structure(self, data: Dict[str, Any], question: str, context: str = "") -> Dict[str, Any]:
        """
        Use LLM to intelligently decide what metrics, sheets, and analysis are needed
        based on the data characteristics and user question.
        """
        logger.debug(f"Starting LLM analysis for question: {question}")
        
        # Build comprehensive data context for LLM
        data_context = self._build_data_context(data)
        logger.debug(f"Built data context: {len(data_context)} characters")
        
        system_prompt = """
        You are an expert data analyst and Excel specialist. Analyze the provided data and user question to determine:
        1. What specific metrics should be calculated
        2. How many analysis sheets are needed and their purposes
        3. Appropriate sheet names (max 15 characters, Excel-safe)
        4. Most relevant analysis types for this data

        Return a JSON structure with:
        {
            "analysis_plan": {
                "total_sheets": number,
                "analysis_focus": "brief description",
                "key_insights_to_find": ["insight1", "insight2"]
            },
            "sheets": [
                {
                    "name": "SheetName",
                    "purpose": "Clear description",
                    "metrics": [
                        {
                            "name": "Human readable name",
                            "formula_type": "SUM|AVERAGE|COUNT|MAX|MIN|MEDIAN|STDEV|VAR|Q1|Q3|IQR|RANGE|SKEW|CV|UNIQUE_COUNT|POSITIVE_COUNT|etc",
                            "target_column": "column_name",
                            "description": "What this tells us",
                            "filter_condition": "optional filter for SUMIF/COUNTIF"
                        }
                    ],
                    "priority": "high|medium|low"
                }
            ]
        }
        """
        
        user_prompt = f"""
        DATA CHARACTERISTICS:
        {data_context}
        
        USER QUESTION: {question}
        
        ADDITIONAL CONTEXT: {context}
        
        Please analyze and recommend the optimal Excel analysis structure. Focus on actionable insights and practical metrics.
        """
        
        try:
            result = self._call_llm_for_structure(system_prompt, user_prompt)
            logger.info(f"LLM analysis completed successfully. Recommended {len(result.get('sheets', []))} sheets")
            return result
        except Exception as e:
            logger.error(f"LLM analysis failed: {str(e)}")
            return self._fallback_intelligent_structure(data, question)

    def _build_data_context(self, data: Dict[str, Any]) -> str:
        """Build comprehensive context about the data for LLM analysis"""
        logger.debug("Building comprehensive data context for LLM")
        
        context_parts = []
        
        # Basic data info
        sample_size = len(data.get("pattern_sample", []))
        total_size = data.get("original_size", 0)
        context_parts.append(f"Dataset size: {total_size:,} total rows, {sample_size} sample rows")
        
        # Column analysis
        ha = data.get("header_analysis", {})
        if ha:
            numeric_cols = [k for k, v in ha.items() if isinstance(v, dict) and v.get("is_numeric")]
            categorical_cols = [k for k, v in ha.items() if isinstance(v, dict) and v.get("is_categorical")]
            date_cols = [k for k, v in ha.items() if isinstance(v, dict) and v.get("is_date")]
            
            context_parts.append(f"Numeric columns ({len(numeric_cols)}): {numeric_cols}")
            context_parts.append(f"Categorical columns ({len(categorical_cols)}): {categorical_cols}")
            if date_cols:
                context_parts.append(f"Date columns ({len(date_cols)}): {date_cols}")
            
            logger.debug(f"Column analysis: {len(numeric_cols)} numeric, {len(categorical_cols)} categorical, {len(date_cols)} date")
        
        # Sample data preview
        ps = data.get("pattern_sample", [])
        print(f"Pattern sample: {len(ps)}")
        if ps and len(ps) > 1:
            headers = ps[0] if isinstance(ps[0], list) else []
            sample_rows = ps[1:4] if len(ps) > 1 else []
            if headers:
                context_parts.append(f"Column headers: {headers}")
                context_parts.append("Sample data (first 3 rows):")
                for i, row in enumerate(sample_rows, 1):
                    # Truncate long values for context
                    truncated_row = [str(val)[:50] + "..." if len(str(val)) > 50 else val for val in row]
                    context_parts.append(f"  Row {i}: {truncated_row}")
        
        # Data quality insights
        context_parts.append(f"Sampling strategy: {data.get('sampling_strategy', 'unknown')}")
        
        full_context = "\n".join(context_parts)
        logger.debug(f"Data context built: {len(full_context)} characters, {len(context_parts)} sections")
        return full_context

    def _call_llm_for_structure(self, system_prompt: str, user_prompt: str) -> Dict[str, Any]:
        """Call LLM to get intelligent analysis structure"""
        try:
            from llm_config import get_openai_client
            logger.debug("Initializing OpenAI client for structure generation")
            client = get_openai_client()
            
            logger.info("Calling OpenAI API for intelligent worksheet structure")
            response = client.chat.completions.create(
                model=get_model(),
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            logger.info("Successfully received LLM response for worksheet structure")
            logger.debug(f"LLM response: {json.dumps(result, indent=2)}")
            
            return self._validate_and_enhance_llm_response(result)
            
        except ImportError:
            logger.error("OpenAI library not installed")
            return self._fallback_intelligent_structure()
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse LLM JSON response: {str(e)}")
            return self._fallback_intelligent_structure()
        except Exception as e:
            logger.error(f"LLM structure generation failed: {str(e)}")
            return self._fallback_intelligent_structure()

    def _validate_and_enhance_llm_response(self, llm_response: Dict[str, Any]) -> Dict[str, Any]:
        """Validate and enhance LLM response with safety checks"""
        logger.debug("Validating and enhancing LLM response")
        
        # Ensure basic structure exists
        if "sheets" not in llm_response:
            llm_response["sheets"] = []
        if "analysis_plan" not in llm_response:
            llm_response["analysis_plan"] = {"total_sheets": 1, "analysis_focus": "General analysis"}
        
        enhanced_sheets = []
        for i, sheet in enumerate(llm_response.get("sheets", [])):
            logger.debug(f"Processing sheet {i+1}: {sheet.get('name', 'unnamed')}")
            
            # Ensure sheet names are Excel-safe
            original_name = sheet.get("name", f"Analysis_{i+1}")
            safe_name = self.generate_sheet_name(original_name)
            sheet["name"] = safe_name
            
            if safe_name != original_name:
                logger.debug(f"Sheet name changed from '{original_name}' to '{safe_name}' for Excel compatibility")
            
            # Validate and enhance metrics
            enhanced_metrics = []
            for j, metric in enumerate(sheet.get("metrics", [])):
                if not metric.get("formula_type"):
                    metric["formula_type"] = "SUM"  # Safe default
                
                if not metric.get("target_column"):
                    metric["target_column"] = "A"  # Safe default
                    logger.debug(f"Added default target_column 'A' to metric {j+1}")
                
                if not metric.get("name"):
                    metric["name"] = f"Metric {j+1}"
                    logger.debug(f"Added default name to metric {j+1}")
                
                enhanced_metrics.append(metric)
            
            sheet["metrics"] = enhanced_metrics
            enhanced_sheets.append(sheet)
        
        llm_response["sheets"] = enhanced_sheets
        logger.info(f"LLM response validated and enhanced: {len(enhanced_sheets)} sheets processed")
        return llm_response

    def _fallback_intelligent_structure(self, data: Dict[str, Any] = None, question: str = "") -> Dict[str, Any]:
        """Intelligent fallback when LLM is unavailable"""
        logger.info("Using intelligent fallback structure generation")
        
        if not data:
            return {
                "analysis_plan": {
                    "total_sheets": 1,
                    "analysis_focus": "Basic analysis",
                    "key_insights_to_find": ["Data overview"]
                },
                "sheets": [{
                    "name": "Analysis",
                    "purpose": "Basic data analysis",
                    "metrics": [],
                    "priority": "high"
                }]
            }
        
        # Analyze data to create intelligent fallback
        ha = data.get("header_analysis", {})
        numeric_cols = [k for k, v in ha.items() if isinstance(v, dict) and v.get("is_numeric")]
        categorical_cols = [k for k, v in ha.items() if isinstance(v, dict) and v.get("is_categorical")]
        
        sheets = []
        
        # Create summary sheet
        summary_metrics = []
        if numeric_cols:
            for col in numeric_cols[:3]:  # Limit to first 3 numeric columns
                summary_metrics.extend([
                    {"name": f"Total {col}", "formula_type": "SUM", "target_column": col, "description": f"Sum of all {col} values"},
                    {"name": f"Average {col}", "formula_type": "AVERAGE", "target_column": col, "description": f"Average {col} value"}
                ])
        
        sheets.append({
            "name": "Summary",
            "purpose": "Key metrics and totals",
            "metrics": summary_metrics,
            "priority": "high"
        })
        
        # Create breakdown sheet if categorical data exists
        if categorical_cols and numeric_cols:
            breakdown_metrics = []
            cat_col = categorical_cols[0]
            num_col = numeric_cols[0]
            breakdown_metrics.extend([
                {"name": f"Unique {cat_col}", "formula_type": "COUNTA", "target_column": cat_col, "description": f"Count of {cat_col} entries"},
                {"name": f"Sum by {cat_col}", "formula_type": "SUMIF", "target_column": num_col, "description": f"Total {num_col} grouped by {cat_col}"}
            ])
            
            sheets.append({
                "name": "Breakdown",
                "purpose": f"Analysis by {cat_col}",
                "metrics": breakdown_metrics,
                "priority": "medium"
            })
        
        logger.info(f"Fallback structure created with {len(sheets)} sheets")
        return {
            "analysis_plan": {
                "total_sheets": len(sheets),
                "analysis_focus": "Automated data analysis",
                "key_insights_to_find": ["Summary statistics", "Data patterns"]
            },
            "sheets": sheets
        }

    # =========================
    # LLM-Driven Workbook Creation
    # =========================

    def create_llm_driven_workbook(self, data: Dict[str, Any], question: str, context: str = "") -> Dict[str, Any]:
        """Create workbook structure based on LLM analysis"""
        logger.info(f"Creating LLM-driven workbook for question: {question}")
        
        try:
            # Get LLM recommendations
            llm_structure = self.llm_decide_metrics_and_structure(data, question, context)
            
            sheets = []
            header_map = self.map_headers_to_letters(data)
            logger.debug(f"Header mapping created: {header_map}")
            
            # Create sheets based on LLM recommendations
            for i, sheet_plan in enumerate(llm_structure.get("sheets", []), 1):
                logger.info(f"Creating sheet {i}: {sheet_plan.get('name', 'unnamed')}")
                sheet = self._create_sheet_from_llm_plan(sheet_plan, data, header_map)
                sheets.append(sheet)
            
            # Add summary sheet with LLM insights
            logger.info("Creating LLM summary sheet")
            summary_sheet = self._create_llm_summary_sheet(llm_structure, question)
            sheets.insert(0, summary_sheet)
            
            result = {
                "sheets": sheets,
                "analysis_plan": llm_structure.get("analysis_plan", {}),
                "metadata": {
                    "created": datetime.now().isoformat(),
                    "llm_driven": True,
                    "total_recommended_sheets": len(sheets),
                    "analysis_focus": llm_structure.get("analysis_plan", {}).get("analysis_focus", ""),
                    "question": question,
                    "context": context
                }
            }
            
            logger.info(f"LLM-driven workbook created successfully with {len(sheets)} sheets")
            return result
            
        except Exception as e:
            logger.error(f"Failed to create LLM-driven workbook: {str(e)}")
            return self.create_fallback_structure(data, question)

    def _create_sheet_from_llm_plan(self, sheet_plan: Dict[str, Any], data: Dict[str, Any], header_map: Dict[str, str], source_sheet: str = "Sheet1") -> Dict[str, Any]:
        """Create individual sheet based on LLM recommendations"""
        sheet_name = sheet_plan.get("name", "Analysis")
        #logger.debug(f"Creating sheet '{sheet_name}' with {len(sheet_plan.get('metrics', []))} metrics")
        
        formulas = []
        cells = []
        
        # Add title and headers
        cells.extend([
            {"address": "A1", "value": sheet_plan.get("purpose", "Analysis")},
            {"address": "A2", "value": "Metric"},
            {"address": "B2", "value": "Value"},
            {"address": "C2", "value": "Description"}
        ])
        
        row = 3
        # Generate formulas based on LLM metrics
        for metric in sheet_plan.get("metrics", []):
            metric_name = metric.get("name", "Unknown Metric")
            #logger.debug(f"Processing metric: {metric_name}")
            
            # Human readable label
            cells.append({
                "address": f"A{row}",
                "value": metric_name
            })
            
            # Generate appropriate formula
            try:
                formula = self._generate_formula_from_metric(metric, header_map, source_sheet)
                formulas.append({
                    "address": f"B{row}",
                    "formula": formula,
                    "description": metric.get("description", "")
                })

            except Exception as e:
                logger.error(f"Failed to generate formula for metric {metric_name}: {str(e)}")
                # Add fallback formula
                formulas.append({
                    "address": f"B{row}",
                    "formula": f"=SUM({source_sheet}!A:A)",
                    "description": "Fallback formula"
                })
            
            # Add description
            cells.append({
                "address": f"C{row}",
                "value": metric.get("description", "")
            })
            
            row += 1
        
        result = {
            "name": sheet_name,
            "type": "llm_generated",
            "purpose": sheet_plan.get("purpose"),
            "priority": sheet_plan.get("priority", "medium"),
            "cells": cells,
            "formulas": formulas,
            "llm_metrics": sheet_plan.get("metrics", [])
        }

        return result

    def _generate_formula_from_metric(self, metric: Dict[str, Any], header_map: Dict[str, str], 
                                     source_sheet: str = "Sheet1", data_context: Dict[str, Any] = None, 
                                     query_context: str = "") -> str:
        """
        Generate Excel formula based on metric specification with optional LLM enhancement.
        
        Args:
            metric: Metric specification dictionary
            header_map: Column name to letter mapping
            source_sheet: Source worksheet name
            data_context: Optional data context for LLM generation
            query_context: Optional user query for LLM generation
            
        Returns:
            Generated Excel formula string
        """
        formula_type = metric.get("formula_type", "SUM").upper()
        target_col = metric.get("target_column", "A")
        filter_condition = metric.get("filter_condition", "")
        use_llm = metric.get("use_llm_generation", False)
        
        # Try LLM generation for complex cases or when explicitly requested
        if (use_llm or self._should_use_llm_generation(metric, query_context)) and data_context and self.openai_client:
            try:

                llm_formula = self._generate_llm_formula(
                    query_context=query_context or f"Generate {formula_type} for {target_col}",
                    data_context=data_context,
                    metric_requirements=metric,
                    source_sheet=source_sheet
                )
                return llm_formula
            except Exception as e:
                logger.warning(f"LLM formula generation failed, falling back to template: {str(e)}")
        
        # Convert column name to letter if mapping exists
        original_col = target_col
        if header_map and target_col in header_map:
            target_col = header_map[target_col]
            logger.debug(f"Column mapped: {original_col} -> {target_col}")
        
        # Generate appropriate formula based on type with dynamic sheet references
        formula_templates = {
            # Basic Aggregations
            "SUM": f"=SUM({source_sheet}!{target_col}:{target_col})",
            "AVERAGE": f"=AVERAGE({source_sheet}!{target_col}:{target_col})",
            "COUNT": f"=COUNT({source_sheet}!{target_col}:{target_col})",
            "COUNTA": f"=COUNTA({source_sheet}!{target_col}:{target_col})",
            "MAX": f"=MAX({source_sheet}!{target_col}:{target_col})",
            "MIN": f"=MIN({source_sheet}!{target_col}:{target_col})",
            
            # Alternative Names
            "TOTAL": f"=SUM({source_sheet}!{target_col}:{target_col})",
            "MEAN": f"=AVERAGE({source_sheet}!{target_col}:{target_col})",
            "AVG": f"=AVERAGE({source_sheet}!{target_col}:{target_col})",
            "MAXIMUM": f"=MAX({source_sheet}!{target_col}:{target_col})",
            "MINIMUM": f"=MIN({source_sheet}!{target_col}:{target_col})",
            "FREQUENCY": f"=COUNT({source_sheet}!{target_col}:{target_col})",
            
            # Statistical Functions
            "MEDIAN": f"=MEDIAN({source_sheet}!{target_col}:{target_col})",
            "STDEV": f"=STDEV.P({source_sheet}!{target_col}:{target_col})",
            "STD": f"=STDEV.P({source_sheet}!{target_col}:{target_col})",
            "STANDARD_DEVIATION": f"=STDEV.P({source_sheet}!{target_col}:{target_col})",
            "VAR": f"=VAR.P({source_sheet}!{target_col}:{target_col})",
            "VARIANCE": f"=VAR.P({source_sheet}!{target_col}:{target_col})",
            "MODE": f"=MODE.SNGL({source_sheet}!{target_col}:{target_col})",
            
            # Percentiles and Quantiles
            "Q1": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},1)",
            "FIRST_QUARTILE": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},1)",
            "Q3": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},3)",
            "THIRD_QUARTILE": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},3)",
            "P90": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.9)",
            "90TH_PERCENTILE": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.9)",
            "P95": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.95)",
            "95TH_PERCENTILE": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.95)",
            "P99": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.99)",
            "99TH_PERCENTILE": f"=PERCENTILE.INC({source_sheet}!{target_col}:{target_col},0.99)",
            
            # Range and Spread  
            "RANGE": f"=MAX({source_sheet}!{target_col}:{target_col})-MIN({source_sheet}!{target_col}:{target_col})",
            "IQR": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},3)-QUARTILE.INC({source_sheet}!{target_col}:{target_col},1)",
            "INTERQUARTILE_RANGE": f"=QUARTILE.INC({source_sheet}!{target_col}:{target_col},3)-QUARTILE.INC({source_sheet}!{target_col}:{target_col},1)",
            "UNIQUE_COUNT": f"=SUMPRODUCT(1/COUNTIF({source_sheet}!{target_col}:{target_col},{source_sheet}!{target_col}:{target_col}))",
            "DISTINCT_COUNT": f"=SUMPRODUCT(1/COUNTIF({source_sheet}!{target_col}:{target_col},{source_sheet}!{target_col}:{target_col}))",
            
            # Business Analytics
            "CV": f"=STDEV.P({source_sheet}!{target_col}:{target_col})/AVERAGE({source_sheet}!{target_col}:{target_col})",
            "COEFFICIENT_OF_VARIATION": f"=STDEV.P({source_sheet}!{target_col}:{target_col})/AVERAGE({source_sheet}!{target_col}:{target_col})",
            "SKEW": f"=SKEW({source_sheet}!{target_col}:{target_col})",
            "SKEWNESS": f"=SKEW({source_sheet}!{target_col}:{target_col})",
            "KURT": f"=KURT({source_sheet}!{target_col}:{target_col})",
            "KURTOSIS": f"=KURT({source_sheet}!{target_col}:{target_col})",
            
            # Growth and Change
            "FIRST": f"=INDEX({source_sheet}!{target_col}:{target_col},1)",
            "EARLIEST": f"=INDEX({source_sheet}!{target_col}:{target_col},1)",
            "LAST": f"=INDEX({source_sheet}!{target_col}:{target_col},COUNTA({source_sheet}!{target_col}:{target_col}))",
            "LATEST": f"=INDEX({source_sheet}!{target_col}:{target_col},COUNTA({source_sheet}!{target_col}:{target_col}))",
            
            # Custom Counts
            "POSITIVE_COUNT": f"=COUNTIF({source_sheet}!{target_col}:{target_col},\">0\")",
            "NEGATIVE_COUNT": f"=COUNTIF({source_sheet}!{target_col}:{target_col},\"<0\")",
            "ZERO_COUNT": f"=COUNTIF({source_sheet}!{target_col}:{target_col},\"=0\")",
            "NON_ZERO_COUNT": f"=COUNTIF({source_sheet}!{target_col}:{target_col},\"<>0\")",
            
            # Conditional Functions (existing)
            "COUNTIF": f"=COUNTIF({source_sheet}!{target_col}:{target_col},\"{filter_condition or '>0'}\")",
            "SUMIF": f"=SUMIF({source_sheet}!{target_col}:{target_col},\"{filter_condition or '>0'}\",{source_sheet}!{target_col}:{target_col})",
            "RANK": f"=RANK.EQ(MAX({source_sheet}!{target_col}:{target_col}),{source_sheet}!{target_col}:{target_col},0)"
        }
        
        formula = formula_templates.get(formula_type, f"=SUM({source_sheet}!{target_col}:{target_col})")
        return formula

    def _generate_llm_formula(self, query_context: str, data_context: Dict[str, Any], 
                             metric_requirements: Dict[str, Any], source_sheet: str = "Sheet1") -> str:
        """
        Generate custom Excel formulas using LLM based on query context and data structure.
        
        Args:
            query_context: User's natural language query
            data_context: Information about data structure (columns, types, sample data)
            metric_requirements: Specific requirements (aggregation type, target columns, etc.)
            source_sheet: Source worksheet name for formula references
            
        Returns:
            Generated Excel formula string
        """
        try:
            if not self.openai_client:
                logger.warning("OpenAI client not available, falling back to template-based generation")
                return self._fallback_to_template_formula(metric_requirements, source_sheet)
            
            # Prepare context for LLM
            context = self._prepare_formula_context(query_context, data_context, metric_requirements, source_sheet)
            
            # Generate formula using LLM
            formula = self._call_llm_for_formula(context)
            
            # Validate and clean the formula
            validated_formula = self._validate_and_clean_formula(formula, data_context, source_sheet)
            
            return validated_formula
            
        except Exception as e:
            logger.error(f"LLM formula generation failed: {str(e)}, falling back to template")
            return self._fallback_to_template_formula(metric_requirements, source_sheet)
    
    def _prepare_formula_context(self, query_context: str, data_context: Dict[str, Any], 
                                metric_requirements: Dict[str, Any], source_sheet: str) -> Dict[str, Any]:
        """Prepare comprehensive context for LLM formula generation."""
        
        # Extract available columns and their types
        header_analysis = data_context.get("header_analysis", {})
        pattern_sample = data_context.get("pattern_sample", [])
        
        columns_info = []
        if pattern_sample and len(pattern_sample) > 0:
            headers = pattern_sample[0] if isinstance(pattern_sample[0], list) else []
            for i, header in enumerate(headers):
                col_letter = _to_col_letter(i)
                col_info = header_analysis.get(header, {})
                columns_info.append({
                    "name": header,
                    "letter": col_letter,
                    "type": "numeric" if col_info.get("is_numeric", False) else "text",
                    "sample_values": col_info.get("sample_values", [])[:3]  # First 3 sample values
                })
        
        # Sample data for better context
        sample_data = []
        if len(pattern_sample) > 1:
            sample_data = pattern_sample[1:6]  # First 5 data rows
        
        context = {
            "user_query": query_context,
            "source_sheet": source_sheet,
            "columns": columns_info,
            "sample_data": sample_data,
            "metric_requirements": metric_requirements,
            "data_size": len(pattern_sample) - 1 if pattern_sample else 0
        }
        
        return context
    
    def _call_llm_for_formula(self, context: Dict[str, Any]) -> str:
        """
        Enhanced OpenAI API call to generate sophisticated Excel formulas.
        
        Recent fix: Improved LLM formula generation with:
        - Enhanced system prompt for complex aggregations
        - Better error handling and validation
        - Support for multi-step calculations
        - Dynamic formula adaptation based on data context
        """
        
        system_prompt = """You are an expert Excel formula architect specializing in complex business analytics formulas. Generate sophisticated, production-ready Excel formulas.

ENHANCED CAPABILITIES:
1. Complex multi-criteria aggregations (SUMIFS, COUNTIFS, AVERAGEIFS with multiple conditions)
2. Advanced statistical functions (PERCENTILE.INC, QUARTILE.INC, STDEV.P, VAR.P)
3. Array formulas and dynamic calculations (SUMPRODUCT, IF arrays, nested functions)
4. Time-based calculations (DATE functions, period comparisons, rolling averages)
5. Business logic formulas (profit margins, growth rates, ratios, KPIs)
6. Nested conditional logic (nested IFs, CHOOSE, SWITCH functions)
7. Lookup and reference formulas (INDEX/MATCH, XLOOKUP, dynamic references)
8. Text manipulation and parsing (CONCATENATE, TEXT functions, regex-like operations)

CRITICAL FORMULA REQUIREMENTS:
- Return ONLY the Excel formula (starting with =)
- Use proper Excel function syntax with correct parentheses and commas
- Reference columns by letters (A, B, C) with proper sheet references (SheetName!A:A)
- Use full column references (A:A) for dynamic data ranges
- Handle empty cells and errors with IFERROR when appropriate
- Optimize for performance with efficient function choices

ADVANCED EXCEL FUNCTIONS AVAILABLE:
Basic Aggregations: SUM, AVERAGE, COUNT, COUNTA, MAX, MIN, MEDIAN, MODE.SNGL
Conditional: SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS
Statistical: STDEV.P, STDEV.S, VAR.P, VAR.S, QUARTILE.INC, PERCENTILE.INC, RANK.EQ
Array: SUMPRODUCT, IF (array), FILTER, SORT, UNIQUE (if Office 365)
Lookup: INDEX, MATCH, VLOOKUP, HLOOKUP, XLOOKUP, OFFSET
Logical: IF, AND, OR, NOT, IFS, SWITCH, CHOOSE
Date/Time: DATE, DATEDIF, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, NETWORKDAYS
Text: CONCATENATE, TEXTJOIN, LEFT, RIGHT, MID, FIND, SEARCH, SUBSTITUTE, TRIM
Math: ROUND, ROUNDUP, ROUNDDOWN, ABS, POWER, SQRT, MOD
Financial: NPV, IRR, PMT, FV, PV (if applicable)

COMPLEX FORMULA PATTERNS:
Multi-criteria aggregation: =SUMIFS(Sheet1!C:C,Sheet1!A:A,"criteria1",Sheet1!B:B,">="&D1,Sheet1!B:B,"<="&E1)
Array calculations: =SUMPRODUCT((Sheet1!A:A="criteria")*(Sheet1!B:B>0)*Sheet1!C:C)
Conditional statistics: =AVERAGE(IF((Sheet1!A:A="criteria")*(Sheet1!B:B>0),Sheet1!C:C))
Dynamic percentiles: =PERCENTILE.INC(IF(Sheet1!A:A="criteria",Sheet1!B:B),0.9)
Growth calculations: =(INDEX(Sheet1!A:A,COUNTA(Sheet1!A:A))-INDEX(Sheet1!A:A,2))/INDEX(Sheet1!A:A,2)
Ratio formulas: =IFERROR(SUM(Sheet1!B:B)/SUM(Sheet1!C:C),0)

BUSINESS INTELLIGENCE FORMULAS:
KPI Calculations: =IFERROR((SUM(Sheet1!Revenue:Revenue)-SUM(Sheet1!Cost:Cost))/SUM(Sheet1!Revenue:Revenue),0)
Year-over-year growth: =(SUMIFS(Sheet1!B:B,Sheet1!A:A,YEAR(TODAY()))-SUMIFS(Sheet1!B:B,Sheet1!A:A,YEAR(TODAY())-1))/SUMIFS(Sheet1!B:B,Sheet1!A:A,YEAR(TODAY())-1)
Moving averages: =AVERAGE(OFFSET(Sheet1!A1,-6,0,7,1))
Ranking with ties: =RANK.EQ(B1,Sheet1!B:B,0)

ERROR HANDLING BEST PRACTICES:
- Wrap complex formulas in IFERROR to handle division by zero, #N/A, etc.
- Use IFNA for lookup functions specifically
- Consider edge cases like empty datasets or invalid criteria
- Provide meaningful fallback values (0, "", or descriptive text)

OPTIMIZATION PRINCIPLES:
- Use SUMIFS instead of SUMPRODUCT when possible for better performance
- Minimize volatile functions (INDIRECT, OFFSET) unless necessary
- Prefer INDEX/MATCH over VLOOKUP for flexibility and speed
- Use structured references when working with Excel tables"""

        # Enhanced user prompt with more context
        complexity_analysis = self._analyze_formula_complexity(context)
        
        user_prompt = f"""Generate an advanced Excel formula for this analytical request:

BUSINESS QUERY: "{context['user_query']}"

DATA ARCHITECTURE:
{self._format_enhanced_columns_for_prompt(context['columns'])}

ANALYTICAL REQUIREMENTS:
{self._format_enhanced_requirements_for_prompt(context['metric_requirements'])}

SAMPLE DATA STRUCTURE:
{self._format_enhanced_sample_data_for_prompt(context['sample_data'], context['columns'])}

COMPLEXITY ANALYSIS:
{complexity_analysis}

DATA INSIGHTS:
- Total data size: {context.get('data_size', 'Unknown')} rows
- Sheet reference: {context['source_sheet']}
- Expected output: Single formula result for analytical dashboard

FORMULA GENERATION INSTRUCTIONS:
1. Analyze the business query to understand the analytical intent
2. Identify the most appropriate Excel functions for the calculation
3. Handle multiple criteria and complex conditions efficiently
4. Include error handling for robust production use
5. Optimize for performance with large datasets
6. Ensure the formula is self-contained and doesn't require helper columns

Generate the most sophisticated and accurate Excel formula that solves this analytical challenge.
Return ONLY the complete formula (starting with =):"""

        try:
            # Use requested model for complex formula generation
            response = self.openai_client.chat.completions.create(
                model=get_model(),  # Upgraded model for better formula generation
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                max_tokens=800,  # Increased for complex formulas
                temperature=0.05,  # Very low temperature for maximum precision
                presence_penalty=0.1,  # Slight penalty to avoid repetitive patterns
                frequency_penalty=0.1   # Encourage diverse function usage
            )
            
            formula = response.choices[0].message.content.strip()
            return formula
            
        except Exception as e:
            logger.error(f"Enhanced OpenAI API call failed: {str(e)}")
            # Try fallback with simpler model
            try:
                logger.info("Attempting fallback with %s", get_model())
                response = self.openai_client.chat.completions.create(
                    model=get_model(),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    max_tokens=500,
                    temperature=0.1
                )
                formula = response.choices[0].message.content.strip()
                return formula
            except Exception as fallback_error:
                logger.error(f"Fallback API call also failed: {str(fallback_error)}")
                raise e  # Raise original error
    
    def _analyze_formula_complexity(self, context: Dict[str, Any]) -> str:
        """Analyze the complexity of the formula requirements for enhanced LLM generation."""
        query = context.get('user_query', '').lower()
        requirements = context.get('metric_requirements', {})
        
        complexity_indicators = []
        
        # Multi-criteria detection
        if any(word in query for word in ['and', 'where', 'multiple', 'both']):
            complexity_indicators.append("Multi-criteria analysis detected")
        
        # Time-based analysis
        if any(word in query for word in ['date', 'time', 'period', 'month', 'year', 'trend']):
            complexity_indicators.append("Time-based calculations required")
        
        # Statistical complexity
        if any(word in query for word in ['percentile', 'quartile', 'median', 'deviation', 'variance']):
            complexity_indicators.append("Advanced statistical functions needed")
        
        # Business logic
        if any(word in query for word in ['ratio', 'margin', 'growth', 'rate', 'percentage']):
            complexity_indicators.append("Business logic calculations required")
        
        # Conditional logic
        if any(word in query for word in ['if', 'when', 'conditional', 'unless']):
            complexity_indicators.append("Conditional logic required")
        
        # Array operations
        if any(word in query for word in ['unique', 'distinct', 'filter', 'dynamic']):
            complexity_indicators.append("Array operations may be needed")
        
        if not complexity_indicators:
            complexity_indicators.append("Standard aggregation - templates may suffice")
        
        return "Formula complexity: " + " | ".join(complexity_indicators)
    
    def _format_enhanced_columns_for_prompt(self, columns: List[Dict[str, Any]]) -> str:
        """Enhanced column formatting with data type analysis and business context."""
        if not columns:
            return "No column information available"
        
        formatted = []
        for col in columns:
            # Enhanced column information
            col_info = f"Column {col['letter']}: '{col['name']}'"
            
            # Data type with inference
            data_type = col['type']
            if data_type == 'numeric':
                col_info += " (Numeric - suitable for aggregations, calculations)"
            elif data_type == 'text':
                col_info += " (Text/Categorical - suitable for grouping, criteria)"
            else:
                col_info += f" ({data_type})"
            
            # Sample values with analysis
            if col['sample_values']:
                samples = [str(v) for v in col['sample_values'][:3]]
                col_info += f"\n    Sample values: {', '.join(samples)}"
                
                # Infer business context from column name
                col_name_lower = col['name'].lower()
                if any(word in col_name_lower for word in ['date', 'time']):
                    col_info += " [Temporal dimension]"
                elif any(word in col_name_lower for word in ['amount', 'price', 'cost', 'revenue', 'sales']):
                    col_info += " [Financial measure]"
                elif any(word in col_name_lower for word in ['count', 'quantity', 'number']):
                    col_info += " [Quantity measure]"
                elif any(word in col_name_lower for word in ['region', 'category', 'type', 'status']):
                    col_info += " [Grouping dimension]"
            
            formatted.append(col_info)
        
        return "\n".join(formatted)
    
    def _format_enhanced_requirements_for_prompt(self, requirements: Dict[str, Any]) -> str:
        """Enhanced requirements formatting with analytical context."""
        if not requirements:
            return "No specific requirements - infer from business query"
        
        formatted = []
        
        # Core requirements
        for key, value in requirements.items():
            if value is not None and key != 'use_llm_generation':
                if key == 'formula_type':
                    formatted.append(f"Primary operation: {value} (but consider if more complex aggregation is needed)")
                elif key == 'target_column':
                    formatted.append(f"Target column: {value} (main column for calculation)")
                elif key == 'filter_condition':
                    formatted.append(f"Filter criteria: {value} (conditional requirements)")
                else:
                    formatted.append(f"{key}: {value}")
        
        # Add analytical guidance
        formatted.append("\nAnalytical guidance:")
        formatted.append("- Consider if multiple criteria or conditions are implied in the query")
        formatted.append("- Determine if complex business logic is required beyond basic aggregation")
        formatted.append("- Include error handling for production use")
        formatted.append("- Optimize for performance with large datasets")
        
        return "\n".join(formatted) if formatted else "Derive all requirements from the business query context"
    
    def _format_enhanced_sample_data_for_prompt(self, sample_data: List[List[Any]], columns: List[Dict[str, Any]]) -> str:
        """Enhanced sample data formatting with data insights."""
        if not sample_data or not columns:
            return "No sample data available - generate formula based on column structure"
        
        # Create header row with column letters
        headers = [f"{col['letter']}({col['name']})" for col in columns]
        result = ["Data structure preview:"]
        result.append(" | ".join(headers))
        result.append("-" * len(" | ".join(headers)))
        
        # Add data rows with analysis
        for i, row in enumerate(sample_data[:3]):
            formatted_row = []
            for j, col in enumerate(columns):
                value = row[j] if j < len(row) else ""
                # Enhanced value formatting
                str_value = str(value)
                if len(str_value) > 12:
                    str_value = str_value[:9] + "..."
                formatted_row.append(str_value)
            result.append(" | ".join(formatted_row))
        
        # Add data insights
        result.append("\nData insights:")
        
        # Analyze numeric columns
        numeric_cols = [col for col in columns if col['type'] == 'numeric']
        if numeric_cols:
            result.append(f"- {len(numeric_cols)} numeric columns suitable for aggregation")
        
        # Analyze categorical columns
        text_cols = [col for col in columns if col['type'] == 'text']
        if text_cols:
            result.append(f"- {len(text_cols)} categorical columns suitable for grouping/filtering")
        
        # Data size context
        if len(sample_data) >= 3:
            result.append(f"- Sample shows structured data (use full column references for formulas)")
        
        return "\n".join(result)
    
    def _format_columns_for_prompt(self, columns: List[Dict[str, Any]]) -> str:
        """Format column information for the LLM prompt (legacy method for compatibility)."""
        formatted = []
        for col in columns:
            sample_text = f" (samples: {', '.join(map(str, col['sample_values']))})" if col['sample_values'] else ""
            formatted.append(f"- {col['letter']}: {col['name']} ({col['type']}){sample_text}")
        return "\n".join(formatted)
    
    def _format_requirements_for_prompt(self, requirements: Dict[str, Any]) -> str:
        """Format metric requirements for the LLM prompt."""
        formatted = []
        for key, value in requirements.items():
            if value is not None:
                formatted.append(f"- {key}: {value}")
        return "\n".join(formatted) if formatted else "- General aggregation/calculation"
    
    def _format_sample_data_for_prompt(self, sample_data: List[List[Any]], columns: List[Dict[str, Any]]) -> str:
        """Format sample data for the LLM prompt."""
        if not sample_data or not columns:
            return "No sample data available"
        
        # Create header row
        headers = [col['letter'] for col in columns]
        result = [" | ".join(headers)]
        result.append("-" * len(" | ".join(headers)))
        
        # Add data rows (limit to first 3 for brevity)
        for row in sample_data[:3]:
            formatted_row = []
            for i, col in enumerate(columns):
                value = row[i] if i < len(row) else ""
                # Truncate long values
                str_value = str(value)
                if len(str_value) > 15:
                    str_value = str_value[:12] + "..."
                formatted_row.append(str_value)
            result.append(" | ".join(formatted_row))
        
        return "\n".join(result)
    
    def _validate_and_clean_formula(self, formula: str, data_context: Dict[str, Any], source_sheet: str) -> str:
        """
        Enhanced validation and cleaning of LLM-generated formulas.
        
        Recent fix: Improved formula validation with:
        - Advanced syntax checking
        - Smart error correction
        - Excel function validation
        - Performance optimization detection
        """
        try:
            original_formula = formula
            
            # Step 1: Basic cleanup
            formula = self._basic_formula_cleanup(formula)
            
            # Step 2: Advanced syntax validation
            validation_issues = self._advanced_syntax_validation(formula, source_sheet)
            
            # Step 3: Smart error correction
            if validation_issues:
                formula = self._smart_error_correction(formula, validation_issues, data_context, source_sheet)
            
            # Step 4: Performance optimization check
            formula = self._optimize_formula_performance(formula)
            
            # Step 5: Final validation
            final_validation = self._final_formula_validation(formula, source_sheet)
            
            if not final_validation['is_valid']:
                logger.warning(f"Formula validation failed: {final_validation['errors']}")
                formula = self._reconstruct_formula_from_intent(original_formula, data_context, source_sheet)
            
            return formula
            
        except Exception as e:
            logger.error(f"Enhanced formula validation failed: {str(e)}")
            # Last resort fallback
            return self._basic_fallback_validation(original_formula, source_sheet)
    
    def _basic_formula_cleanup(self, formula: str) -> str:
        """Basic formula cleanup and formatting."""
        # Remove markdown and formatting
        formula = formula.replace("```excel", "").replace("```", "").replace("`", "").strip()
        
        # Ensure formula starts with =
        if not formula.startswith("="):
            formula = "=" + formula
        
        # Remove extra whitespace
        formula = " ".join(formula.split())
        
        # Fix common typos in function names
        function_corrections = {
            "SUMIF(": "SUMIFS(",  # Often more appropriate for complex conditions
            "sum(": "SUM(",
            "average(": "AVERAGE(",
            "count(": "COUNT(",
            "if(": "IF(",
        }
        
        for typo, correction in function_corrections.items():
            if typo in formula and correction not in formula:
                formula = formula.replace(typo, correction)
        
        return formula
    
    def _advanced_syntax_validation(self, formula: str, source_sheet: str) -> List[Dict[str, str]]:
        """Advanced syntax validation with detailed error reporting."""
        issues = []
        
        # Check parentheses balance
        open_parens = formula.count("(")
        close_parens = formula.count(")")
        if open_parens != close_parens:
            issues.append({
                "type": "syntax_error",
                "issue": "unbalanced_parentheses",
                "description": f"Open: {open_parens}, Close: {close_parens}",
                "severity": "high"
            })
        
        # Check for sheet references
        if source_sheet not in formula and "Sheet1" not in formula:
            issues.append({
                "type": "reference_error",
                "issue": "missing_sheet_reference",
                "description": f"Formula should reference {source_sheet}",
                "severity": "medium"
            })
        
        # Check for valid Excel functions
        import re
        function_pattern = r'([A-Z][A-Z0-9]*)\('
        functions_used = re.findall(function_pattern, formula)
        
        valid_excel_functions = {
            'SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'MEDIAN', 'MODE',
            'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS',
            'STDEV', 'STDEVP', 'VAR', 'VARP', 'QUARTILE', 'PERCENTILE',
            'INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'OFFSET',
            'IF', 'IFS', 'AND', 'OR', 'NOT', 'IFERROR', 'IFNA',
            'SUMPRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS',
            'DATE', 'YEAR', 'MONTH', 'DAY', 'TODAY', 'NOW'
        }
        
        for func in functions_used:
            if func not in valid_excel_functions:
                issues.append({
                    "type": "function_error",
                    "issue": "unknown_function",
                    "description": f"Function '{func}' may not be valid in Excel",
                    "severity": "medium"
                })
        
        # Check for common syntax errors
        if '""' in formula and '"' in formula:
            issues.append({
                "type": "syntax_error",
                "issue": "quote_mismatch",
                "description": "Potential quote matching issues",
                "severity": "low"
            })
        
        return issues
    
    def _smart_error_correction(self, formula: str, issues: List[Dict[str, str]], 
                              data_context: Dict[str, Any], source_sheet: str) -> str:
        """Smart error correction based on validation issues."""
        corrected_formula = formula
        
        for issue in issues:
            if issue['issue'] == 'unbalanced_parentheses':
                corrected_formula = self._fix_parentheses_balance(corrected_formula)
            
            elif issue['issue'] == 'missing_sheet_reference':
                corrected_formula = self._add_missing_sheet_references(corrected_formula, source_sheet)
            
            elif issue['issue'] == 'unknown_function':
                corrected_formula = self._replace_unknown_functions(corrected_formula)
        
        return corrected_formula
    
    def _fix_parentheses_balance(self, formula: str) -> str:
        """Attempt to fix unbalanced parentheses."""
        open_count = formula.count("(")
        close_count = formula.count(")")
        
        if open_count > close_count:
            # Add missing closing parentheses at the end
            formula += ")" * (open_count - close_count)
        elif close_count > open_count:
            # Remove extra closing parentheses from the end
            extra_closes = close_count - open_count
            for _ in range(extra_closes):
                if formula.endswith(")"):
                    formula = formula[:-1]
        
        return formula
    
    def _add_missing_sheet_references(self, formula: str, source_sheet: str) -> str:
        """Add missing sheet references to column references."""
        import re
        
        # Pattern to match column references without sheet names
        column_pattern = r'(?<![:\w])([A-Z]+:[A-Z]+)(?![\w:])'
        
        def add_sheet_ref(match):
            col_ref = match.group(1)
            return f"{source_sheet}!{col_ref}"
        
        # Only add sheet references if they're missing
        if source_sheet not in formula:
            formula = re.sub(column_pattern, add_sheet_ref, formula)
        
        return formula
    
    def _replace_unknown_functions(self, formula: str) -> str:
        """Replace unknown functions with Excel-compatible alternatives."""
        replacements = {
            'MEAN(': 'AVERAGE(',
            'STD(': 'STDEV.P(',
            'VAR(': 'VAR.P(',
            'UNIQUE(': 'SUMPRODUCT(1/COUNTIF(',  # Approximate replacement
        }
        
        for unknown, replacement in replacements.items():
            formula = formula.replace(unknown, replacement)
        
        return formula
    
    def _optimize_formula_performance(self, formula: str) -> str:
        """Optimize formula for better performance."""
        # Replace SUMPRODUCT with SUMIFS where possible for better performance
        if 'SUMPRODUCT(' in formula and '=' in formula:
            logger.debug("Complex SUMPRODUCT detected - consider manual optimization")
        
        # Prefer SUMIFS over nested SUM(IF()) for better performance
        if 'SUM(IF(' in formula:
            logger.debug("Nested SUM(IF()) detected - could be optimized to SUMIFS")
        
        return formula
    
    def _final_formula_validation(self, formula: str, source_sheet: str) -> Dict[str, Any]:
        """Final comprehensive validation of the formula."""
        errors = []
        warnings = []
        
        # Check basic structure
        if not formula.startswith('='):
            errors.append("Formula must start with '='")
        
        # Check parentheses balance (final check)
        if formula.count('(') != formula.count(')'):
            errors.append("Unbalanced parentheses after correction")
        
        # Check for essential components
        if not any(func in formula.upper() for func in ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF']):
            warnings.append("No recognized aggregation function found")
        
        # Check sheet reference
        if source_sheet in formula or 'Sheet1' in formula:
            # Good - has sheet reference
            pass
        else:
            warnings.append("No sheet reference found")
        
        return {
            'is_valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings,
            'complexity': 'high' if len(formula) > 100 else 'medium' if len(formula) > 50 else 'low'
        }
    
    def _reconstruct_formula_from_intent(self, original_formula: str, data_context: Dict[str, Any], source_sheet: str) -> str:
        """Reconstruct formula from original intent when validation fails."""
        
        # Extract key components from failed formula
        upper_formula = original_formula.upper()
        
        # Determine primary function
        if 'SUM' in upper_formula:
            primary_func = 'SUM'
        elif 'AVERAGE' in upper_formula:
            primary_func = 'AVERAGE'
        elif 'COUNT' in upper_formula:
            primary_func = 'COUNT'
        else:
            primary_func = 'SUM'  # Default
        
        # Use first numeric column as fallback
        pattern_sample = data_context.get('pattern_sample', [])
        if pattern_sample and len(pattern_sample) > 0:
            headers = pattern_sample[0]
            # Find first numeric-looking column
            target_col = 'A'
            for i, header in enumerate(headers):
                if any(word in str(header).lower() for word in ['amount', 'value', 'price', 'cost', 'revenue']):
                    target_col = _to_col_letter(i)
                    break
        else:
            target_col = 'A'
        
        reconstructed = f"={primary_func}({source_sheet}!{target_col}:{target_col})"
        return reconstructed
    
    def _basic_fallback_validation(self, formula: str, source_sheet: str) -> str:
        """Basic fallback validation when all else fails."""
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Basic cleanup
        formula = formula.replace('```', '').strip()
        
        # If all else fails, return a simple SUM formula
        if not any(func in formula.upper() for func in ['SUM', 'AVERAGE', 'COUNT']):
            return f"=SUM({source_sheet}!A:A)"
        
        return formula
    
    def _fallback_to_template_formula(self, metric_requirements: Dict[str, Any], source_sheet: str) -> str:
        """Fallback to template-based formula generation when LLM fails."""
        
        # Extract basic requirements
        formula_type = metric_requirements.get("formula_type", "SUM").upper()
        target_column = metric_requirements.get("target_column", "A")
        
        # Use existing template method
        mock_metric = {
            "formula_type": formula_type,
            "target_column": target_column,
            "filter_condition": metric_requirements.get("filter_condition", "")
        }
        
        return self._generate_formula_from_metric(mock_metric, {}, source_sheet)
    
    def _should_use_llm_generation(self, metric: Dict[str, Any], query_context: str) -> bool:
        """
        Determine if LLM generation should be used based on complexity indicators.
        
        Args:
            metric: Metric specification dictionary
            query_context: User's query context
            
        Returns:
            True if LLM generation is recommended
        """
        # Use LLM for complex cases that templates can't handle well
        complexity_indicators = [
            # Multiple conditions or criteria
            "multiple" in query_context.lower(),
            "and" in query_context.lower() and "where" in query_context.lower(),
            
            # Complex aggregations
            "weighted" in query_context.lower(),
            "conditional" in query_context.lower(),
            "nested" in query_context.lower(),
            
            # Time-based or date calculations
            "date" in query_context.lower(),
            "time" in query_context.lower(),
            "period" in query_context.lower(),
            "monthly" in query_context.lower(),
            "yearly" in query_context.lower(),
            
            # Statistical operations not in templates
            "correlation" in query_context.lower(),
            "regression" in query_context.lower(),
            "trend" in query_context.lower(),
            
            # Custom business logic
            "if" in query_context.lower() and "then" in query_context.lower(),
            "calculate" in query_context.lower() and ("rate" in query_context.lower() or "ratio" in query_context.lower()),
            
            # Multi-step calculations
            "first" in query_context.lower() and "then" in query_context.lower(),
            "step" in query_context.lower(),
            
            # Non-standard formula types
            metric.get("formula_type", "").upper() not in [
                "SUM", "AVERAGE", "COUNT", "MAX", "MIN", "MEDIAN", "STDEV", "VAR",
                "COUNTIF", "SUMIF", "AVERAGEIF", "TOTAL", "MEAN", "AVG"
            ]
        ]
        
        # Use LLM if any complexity indicators are present
        should_use_llm = any(complexity_indicators)
        
        if should_use_llm:
            logger.info(f"LLM generation recommended due to complexity indicators in query: '{query_context}'")
        
        return should_use_llm
    
    # =========================
    # Enhanced LLM Formula Generation - Public API
    # =========================
    
    def generate_dynamic_formula(self, query_context: str, data_context: Dict[str, Any], 
                                formula_requirements: Dict[str, Any] = None, 
                                source_sheet: str = "Sheet1", 
                                force_llm: bool = False) -> Dict[str, Any]:
        """
        Enhanced public method for dynamic LLM-based formula generation.
        
        Recent fix: New enhanced LLM formula generation API with:
        - Dynamic complexity analysis and LLM selection
        - Advanced error handling and smart fallbacks
        - Comprehensive formula validation and optimization
        - Backward compatibility with existing template system
        
        Args:
            query_context: User's natural language query describing what they want to calculate
            data_context: Complete data context with columns, types, samples, header analysis
            formula_requirements: Optional specific requirements (formula_type, target_column, etc.)
            source_sheet: Source worksheet name for formula references
            force_llm: Force LLM usage even for simple cases
            
        Returns:
            Dict containing:
            - formula: Generated Excel formula
            - method_used: "llm_enhanced", "llm_basic", or "template_fallback"
            - complexity_analysis: Analysis of formula complexity
            - validation_results: Validation and error checking results
            - performance_notes: Performance optimization recommendations
            - metadata: Generation metadata including timing and model used
        """
        start_time = datetime.now()
        
        try:
            # Step 1: Prepare requirements if not provided
            if not formula_requirements:
                formula_requirements = self._extract_requirements_from_query(query_context)
            
            # Step 2: Enhanced complexity analysis
            complexity_analysis = self._analyze_enhanced_complexity(query_context, data_context, formula_requirements)
            
            # Step 3: Determine generation method
            should_use_llm = force_llm or complexity_analysis['should_use_llm']
            method_used = None
            formula = None
            validation_results = None
            
            if should_use_llm and self.openai_client:
                try:
                    # Enhanced LLM generation
                    formula = self._generate_enhanced_llm_formula(
                        query_context=query_context,
                        data_context=data_context,
                        metric_requirements=formula_requirements,
                        source_sheet=source_sheet,
                        complexity_info=complexity_analysis
                    )
                    method_used = "llm_enhanced"
                    
                    # Enhanced validation
                    validation_results = self._comprehensive_formula_validation(formula, data_context, source_sheet)
                    
                except Exception as e:
                    logger.warning(f"Enhanced LLM generation failed: {str(e)}, trying basic LLM")
                    try:
                        # Fallback to basic LLM
                        formula = self._generate_llm_formula(
                            query_context=query_context,
                            data_context=data_context,
                            metric_requirements=formula_requirements,
                            source_sheet=source_sheet
                        )
                        method_used = "llm_basic"
                        validation_results = self._comprehensive_formula_validation(formula, data_context, source_sheet)
                        
                    except Exception as e2:
                        logger.warning(f"Basic LLM generation also failed: {str(e2)}, using template fallback")
                        formula = None
            
            # Step 4: Template fallback if LLM failed or not used
            if not formula:
                logger.info("Using enhanced template generation")
                metric_for_template = {
                    **formula_requirements,
                    "use_llm_generation": False
                }
                formula = self._generate_formula_from_metric(
                    metric_for_template, 
                    self.map_headers_to_letters(data_context), 
                    source_sheet, 
                    data_context, 
                    query_context
                )
                method_used = "template_fallback"
                validation_results = self._basic_formula_validation(formula, source_sheet)
            
            # Step 5: Performance analysis
            performance_notes = self._analyze_formula_performance(formula, complexity_analysis)
            
            # Step 6: Generate comprehensive response
            end_time = datetime.now()
            generation_time = (end_time - start_time).total_seconds()
            
            result = {
                "formula": formula,
                "method_used": method_used,
                "complexity_analysis": complexity_analysis,
                "validation_results": validation_results,
                "performance_notes": performance_notes,
                "metadata": {
                    "generation_time_seconds": round(generation_time, 3),
                    "query_length": len(query_context),
                    "data_columns": len(data_context.get('pattern_sample', [[]])[0]) if data_context.get('pattern_sample') else 0,
                    "llm_available": self.openai_client is not None,
                    "forced_llm": force_llm,
                    "timestamp": end_time.isoformat()
                }
            }
            
            return result
            
        except Exception as e:
            logger.error(f"Dynamic formula generation failed completely: {str(e)}")
            # Ultimate fallback
            return {
                "formula": f"=SUM({source_sheet}!A:A)",
                "method_used": "emergency_fallback",
                "complexity_analysis": {"error": str(e)},
                "validation_results": {"is_valid": True, "errors": [], "warnings": ["Emergency fallback used"]},
                "performance_notes": {"note": "Basic SUM formula for safety"},
                "metadata": {
                    "generation_time_seconds": (datetime.now() - start_time).total_seconds(),
                    "error": str(e),
                    "emergency_fallback": True
                }
            }
    
    def _analyze_enhanced_complexity(self, query: str, data_context: Dict[str, Any], 
                                   requirements: Dict[str, Any]) -> Dict[str, Any]:
        """Enhanced complexity analysis for better LLM decision making."""
        query_lower = query.lower()
        
        # Complexity scoring system
        complexity_score = 0
        indicators = []
        
        # Multi-criteria indicators (high complexity)
        multi_criteria_patterns = ['and', 'where', 'with', 'multiple', 'both', 'also']
        multi_criteria_count = sum(1 for pattern in multi_criteria_patterns if pattern in query_lower)
        if multi_criteria_count >= 2:
            complexity_score += 3
            indicators.append(f"Multiple criteria detected ({multi_criteria_count} indicators)")
        
        # Business logic indicators (medium-high complexity)
        business_patterns = ['ratio', 'margin', 'growth', 'rate', 'percentage', 'profit', 'roi', 'kpi']
        business_count = sum(1 for pattern in business_patterns if pattern in query_lower)
        if business_count > 0:
            complexity_score += 2
            indicators.append(f"Business logic required ({business_count} indicators)")
        
        # Statistical indicators (medium complexity)
        stats_patterns = ['percentile', 'quartile', 'median', 'deviation', 'variance', 'correlation']
        stats_count = sum(1 for pattern in stats_patterns if pattern in query_lower)
        if stats_count > 0:
            complexity_score += 2
            indicators.append(f"Statistical functions needed ({stats_count} indicators)")
        
        # Conditional logic indicators (medium complexity)
        conditional_patterns = ['if', 'when', 'unless', 'conditional', 'depending']
        conditional_count = sum(1 for pattern in conditional_patterns if pattern in query_lower)
        if conditional_count > 0:
            complexity_score += 1
            indicators.append(f"Conditional logic required ({conditional_count} indicators)")
        
        # Time-based indicators (medium complexity)
        time_patterns = ['date', 'time', 'period', 'month', 'year', 'trend', 'historical']
        time_count = sum(1 for pattern in time_patterns if pattern in query_lower)
        if time_count > 0:
            complexity_score += 1
            indicators.append(f"Time-based calculations ({time_count} indicators)")
        
        # Data structure complexity
        pattern_sample = data_context.get('pattern_sample', [])
        column_count = len(pattern_sample[0]) if pattern_sample else 0
        if column_count > 10:
            complexity_score += 1
            indicators.append(f"Large dataset ({column_count} columns)")
        
        # Formula type complexity
        formula_type = requirements.get('formula_type', '').upper()
        complex_formula_types = ['PERCENTILE', 'QUARTILE', 'RANK', 'INDEX', 'MATCH']
        if formula_type in complex_formula_types:
            complexity_score += 2
            indicators.append(f"Complex formula type: {formula_type}")
        
        # Decision logic
        should_use_llm = complexity_score >= 2  # Threshold for LLM usage
        
        return {
            "complexity_score": complexity_score,
            "indicators": indicators,
            "should_use_llm": should_use_llm,
            "confidence": min(complexity_score / 5.0, 1.0),  # Normalize to 0-1
            "recommendation": self._get_complexity_recommendation(complexity_score)
        }
    
    def _get_complexity_recommendation(self, score: int) -> str:
        """Get recommendation based on complexity score."""
        if score >= 4:
            return "High complexity - Enhanced LLM with advanced validation recommended"
        elif score >= 2:
            return "Medium complexity - LLM generation recommended"
        else:
            return "Low complexity - Template generation sufficient"
    
    def _generate_enhanced_llm_formula(self, query_context: str, data_context: Dict[str, Any], 
                                     metric_requirements: Dict[str, Any], source_sheet: str,
                                     complexity_info: Dict[str, Any]) -> str:
        """Enhanced LLM formula generation with complexity-aware prompting."""
        
        # Build enhanced context
        enhanced_context = self._prepare_formula_context(query_context, data_context, metric_requirements, source_sheet)
        enhanced_context["complexity_analysis"] = complexity_info
        
        # Use the enhanced LLM call
        formula = self._call_llm_for_formula(enhanced_context)
        
        # Enhanced validation and cleaning
        validated_formula = self._validate_and_clean_formula(formula, data_context, source_sheet)
        
        return validated_formula
    
    def _comprehensive_formula_validation(self, formula: str, data_context: Dict[str, Any], 
                                        source_sheet: str) -> Dict[str, Any]:
        """Comprehensive validation with detailed reporting."""
        try:
            # Use enhanced validation
            basic_validation = self._final_formula_validation(formula, source_sheet)
            
            # Additional checks
            additional_checks = {
                "has_error_handling": "IFERROR" in formula.upper() or "IFNA" in formula.upper(),
                "uses_full_column_refs": ":" in formula,
                "has_sheet_reference": source_sheet in formula or "Sheet1" in formula,
                "formula_length": len(formula),
                "function_count": formula.count("("),
                "nested_functions": formula.count("(") > 1
            }
            
            # Combine results
            return {
                **basic_validation,
                "additional_checks": additional_checks,
                "quality_score": self._calculate_formula_quality_score(formula, basic_validation, additional_checks)
            }
            
        except Exception as e:
            return {
                "is_valid": False,
                "errors": [f"Validation failed: {str(e)}"],
                "warnings": [],
                "additional_checks": {},
                "quality_score": 0
            }
    
    def _calculate_formula_quality_score(self, formula: str, basic_validation: Dict[str, Any], 
                                       additional_checks: Dict[str, Any]) -> float:
        """Calculate a quality score for the formula (0-100)."""
        score = 0
        
        # Basic validity (40 points)
        if basic_validation['is_valid']:
            score += 40
        
        # Error handling (20 points)
        if additional_checks['has_error_handling']:
            score += 20
        
        # Proper references (20 points)
        if additional_checks['uses_full_column_refs']:
            score += 10
        if additional_checks['has_sheet_reference']:
            score += 10
        
        # Complexity appropriateness (20 points)
        complexity = basic_validation.get('complexity', 'low')
        if complexity == 'high' and additional_checks['nested_functions']:
            score += 20
        elif complexity == 'medium':
            score += 15
        elif complexity == 'low':
            score += 10
        
        return min(score, 100)
    
    def _analyze_formula_performance(self, formula: str, complexity_analysis: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze formula performance characteristics."""
        notes = []
        recommendations = []
        
        # Check for performance-intensive functions
        if "SUMPRODUCT" in formula.upper():
            notes.append("SUMPRODUCT detected - may be slow with large datasets")
            recommendations.append("Consider SUMIFS if multiple criteria aggregation")
        
        if "OFFSET" in formula.upper():
            notes.append("OFFSET function detected - volatile function, recalculates frequently")
            recommendations.append("Consider INDEX/MATCH for better performance")
        
        if "INDIRECT" in formula.upper():
            notes.append("INDIRECT function detected - volatile function")
            recommendations.append("Use direct references when possible")
        
        # Check formula complexity vs performance
        if complexity_analysis.get('complexity_score', 0) > 3:
            notes.append("High complexity formula - monitor calculation time")
            recommendations.append("Test with full dataset to ensure acceptable performance")
        
        # Positive performance indicators
        if "SUMIFS" in formula.upper() or "COUNTIFS" in formula.upper():
            notes.append("Using optimized conditional functions")
        
        if "IFERROR" in formula.upper():
            notes.append("Good error handling implemented")
        
        return {
            "performance_notes": notes,
            "recommendations": recommendations,
            "estimated_performance": "good" if len(notes) <= 1 else "monitor" if len(notes) <= 3 else "potential_issues"
        }
    
    def _basic_formula_validation(self, formula: str, source_sheet: str) -> Dict[str, Any]:
        """Basic validation for template-generated formulas."""
        return {
            "is_valid": formula.startswith("=") and formula.count("(") == formula.count(")"),
            "errors": [],
            "warnings": [] if source_sheet in formula else ["No sheet reference found"],
            "complexity": "low",
            "additional_checks": {
                "has_sheet_reference": source_sheet in formula,
                "uses_full_column_refs": ":" in formula,
                "formula_length": len(formula)
            },
            "quality_score": 75 if formula.startswith("=") else 25
        }
    
    def generate_enhanced_formula(self, query: str, data_context: Dict[str, Any], 
                                 formula_requirements: Dict[str, Any] = None, 
                                 source_sheet: str = "Sheet1") -> Dict[str, Any]:
        """
        Public method to generate enhanced formulas with LLM integration.
        
        Args:
            query: User's natural language query
            data_context: Data structure and sample information
            formula_requirements: Optional specific requirements for the formula
            source_sheet: Source worksheet name
            
        Returns:
            Dictionary containing the generated formula and metadata
        """
        try:
            # Prepare requirements if not provided
            if not formula_requirements:
                formula_requirements = self._extract_requirements_from_query(query)
            
            # Add LLM usage flag based on complexity
            formula_requirements["use_llm_generation"] = self._should_use_llm_generation(
                formula_requirements, query
            )
            
            # Generate header mapping
            header_map = self.map_headers_to_letters(data_context)
            
            # Generate formula
            formula = self._generate_formula_from_metric(
                metric=formula_requirements,
                header_map=header_map,
                source_sheet=source_sheet,
                data_context=data_context,
                query_context=query
            )
            
            # Prepare result with metadata
            result = {
                "formula": formula,
                "method_used": "llm" if formula_requirements.get("use_llm_generation") else "template",
                "requirements": formula_requirements,
                "source_sheet": source_sheet,
                "success": True,
                "error": None
            }
            
            return result
            
        except Exception as e:
            logger.error(f"Enhanced formula generation failed: {str(e)}")
            return {
                "formula": f"=SUM({source_sheet}!A:A)",  # Basic fallback
                "method_used": "fallback",
                "requirements": formula_requirements or {},
                "source_sheet": source_sheet,
                "success": False,
                "error": str(e)
            }
    
    def _extract_requirements_from_query(self, query: str) -> Dict[str, Any]:
        """
        Extract formula requirements from natural language query.
        
        Args:
            query: User's natural language query
            
        Returns:
            Dictionary of extracted requirements
        """
        query_lower = query.lower()
        requirements = {}
        
        # Extract aggregation type
        if any(word in query_lower for word in ["sum", "total", "add"]):
            requirements["formula_type"] = "SUM"
        elif any(word in query_lower for word in ["average", "avg", "mean"]):
            requirements["formula_type"] = "AVERAGE"
        elif any(word in query_lower for word in ["count", "number of", "how many"]):
            requirements["formula_type"] = "COUNT"
        elif any(word in query_lower for word in ["max", "maximum", "highest", "largest"]):
            requirements["formula_type"] = "MAX"
        elif any(word in query_lower for word in ["min", "minimum", "lowest", "smallest"]):
            requirements["formula_type"] = "MIN"
        elif any(word in query_lower for word in ["median", "middle"]):
            requirements["formula_type"] = "MEDIAN"
        elif any(word in query_lower for word in ["standard deviation", "stdev", "std"]):
            requirements["formula_type"] = "STDEV"
        else:
            requirements["formula_type"] = "SUM"  # Default
        
        # Extract conditional logic indicators
        if any(phrase in query_lower for phrase in ["where", "if", "when", "for"]):
            requirements["has_conditions"] = True
        
        # Extract column references (simple heuristic)
        words = query.split()
        for i, word in enumerate(words):
            if word.lower() in ["column", "col"] and i + 1 < len(words):
                requirements["target_column"] = words[i + 1].upper()
                break
        
        # Default target column if not found
        if "target_column" not in requirements:
            requirements["target_column"] = "A"
        
        return requirements

    def _create_llm_summary_sheet(self, llm_structure: Dict[str, Any], question: str) -> Dict[str, Any]:
        """Create summary sheet with LLM-generated insights"""
        logger.debug("Creating LLM summary sheet")
        
        analysis_plan = llm_structure.get("analysis_plan", {})
        
        cells = [
            {"address": "A1", "value": "🤖 AI-Generated Analysis Plan"},
            {"address": "A3", "value": "Question:"},
            {"address": "B3", "value": question},
            {"address": "A5", "value": "Analysis Focus:"},
            {"address": "B5", "value": analysis_plan.get('analysis_focus', 'General analysis')},
            {"address": "A7", "value": "Recommended Sheets:"},
            {"address": "B7", "value": str(analysis_plan.get('total_sheets', 1))},
            {"address": "A9", "value": "Key Insights to Discover:"}
        ]
        
        # Add key insights
        insights = analysis_plan.get("key_insights_to_find", [])
        for i, insight in enumerate(insights, 1):
            cells.append({
                "address": f"A{9 + i}",
                "value": f"• {insight}"
            })
        
        # Add sheet navigation
        row = 15
        cells.extend([
            {"address": f"A{row}", "value": "📊 Analysis Sheets:"},
            {"address": f"A{row+1}", "value": "Sheet Name"},
            {"address": f"B{row+1}", "value": "Purpose"},
            {"address": f"C{row+1}", "value": "Priority"}
        ])
        
        for i, sheet in enumerate(llm_structure.get("sheets", []), 2):
            cells.extend([
                {"address": f"A{row + i}", "value": sheet.get("name", "")},
                {"address": f"B{row + i}", "value": sheet.get("purpose", "")},
                {"address": f"C{row + i}", "value": sheet.get('priority', 'medium').title()}
            ])
        
        result = {
            "name": "AI_Overview",
            "type": "llm_summary",
            "cells": cells,
            "formulas": [],
            "analysis_plan": analysis_plan
        }
        
        logger.info("LLM summary sheet created successfully")
        return result

    # =========================
    # Column helpers and safety
    # =========================

    def map_headers_to_letters(self, data: Dict[str, Any]) -> Dict[str, str]:
        """
        Build header name -> column letter mapping from:
        1) pattern_sample first row if it looks like headers,
        2) header_analysis keys order as fallback.
        Returns {} if mapping cannot be inferred.
        """
        logger.debug("Creating header to column letter mapping")
        header_map: Dict[str, str] = {}

        # Preferred: pattern_sample first row as headers
        ps = data.get("pattern_sample") or []
        #print(f"[DEBUG] pattern_sample: {ps}")
        if ps and isinstance(ps, list) and len(ps) > 0 and isinstance(ps[0], list):
            first = ps[0]
            #print(f"[DEBUG] First row (headers): {first}")
            if all(isinstance(x, str) for x in first) and len(set(first)) == len(first):
                for idx, name in enumerate(first):
                    col_letter = _to_col_letter(idx)
                    header_map[name] = col_letter
                    print(f"[DEBUG] Mapping {name} -> {col_letter} (index {idx})")
                print(f"[DEBUG] Final header mapping: {header_map}")
                return header_map

        # Fallback: header_analysis key order
        ha = data.get("header_analysis") or {}
        print(f"[DEBUG] Using fallback header_analysis: {ha}")
        if isinstance(ha, dict) and ha:
            for idx, name in enumerate(ha.keys()):
                col_letter = _to_col_letter(idx)
                header_map[name] = col_letter
                print(f"[DEBUG] Fallback mapping {name} -> {col_letter} (index {idx})")
            print(f"[DEBUG] Final fallback header mapping: {header_map}")
            return header_map

        print("[DEBUG] No header mapping could be created - returning empty dict")
        return header_map

    def convert_name_range_to_letter(self, expr: str, header_map: Dict[str, str]) -> str:
        """
        Convert name-based references like 'Sales:Sales' to 'B:B' using header_map.
        """
        if not header_map or not expr or not isinstance(expr, str):
            return expr

        original_expr = expr
        
        # Replace exact Name:Name occurrences
        def repl_same_colons(m):
            name = m.group(1)
            letter = header_map.get(name)
            return f"{letter}:{letter}" if letter else m.group(0)

        out = re.sub(r'([A-Za-z_][A-Za-z0-9_]*)\s*:\s*\1', repl_same_colons, expr)

        # Handle function-based segments
        for name, letter in header_map.items():
            out = re.sub(rf'([,(])\s*{re.escape(name)}\s*:', rf'\1{letter}:', out)
            out = re.sub(rf':\s*{re.escape(name)}\s*([),])', rf':{letter}\1', out)

        return out

    def normalize_formula_targets(self, formulas: List[Dict[str, Any]], start_col: str = "F", start_row: int = 2) -> List[Dict[str, Any]]:
        """
        Ensure formula write addresses exist and are outside the typical A..D columns used by data.
        """
        
        row = start_row
        out: List[Dict[str, Any]] = []
        for i, f in enumerate(formulas or []):
            addr = f.get("address") or f.get("cell")
            if not addr:
                addr = f"{start_col}{row}"
                row += 1
            
            form = f.get("formula", "")
            if form and not str(form).startswith("="):
                form = "=" + str(form)
            
            out.append({
                "address": addr,
                "formula": form,
                "description": f.get("description", "")
            })
        
        return out

    def generate_llm_sheet_name(self, query: str, analysis_type: str, data_context: Dict[str, Any]) -> str:
        """Generate intelligent sheet names using LLM based on query and data context."""
        try:
            # Extract key context for LLM
            columns = list(data_context.get("header_analysis", {}).keys())[:5]  # Limit to 5 columns
            
            system_prompt = """You are an Excel sheet naming expert. Generate creative, professional sheet names that:
1. Are 8-15 characters long (Excel limit)
2. Clearly indicate the analysis type
3. Use business-friendly language
4. Avoid special characters: : / \\ ? * [ ]
5. Are memorable and descriptive

Examples:
- "Sales_TopN" for top N sales analysis
- "RevBy_Region" for revenue by region
- "Perf_Stats" for performance statistics
- "Trend_2024" for trending analysis"""

            user_prompt = f"""Create a sheet name for this analysis:

Query: "{query}"
Analysis Type: {analysis_type}
Data Columns: {', '.join(columns) if columns else 'various metrics'}

Return only the sheet name, nothing else."""

            response = self.openai_client.chat.completions.create(
                model=DEFAULT_OPENAI_MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                max_tokens=20,
                temperature=0.7
            )
            
            llm_name = response.choices[0].message.content.strip().strip('"').strip("'")
            
            # Clean and validate the LLM suggestion
            return self.clean_sheet_name(llm_name)
            
        except Exception as e:
            logger.warning(f"LLM sheet naming failed: {e}, using fallback")
            return self.clean_sheet_name(analysis_type)
    
    def clean_sheet_name(self, name: str) -> str:
        """Clean and validate sheet name for Excel compatibility."""
        max_len = self.max_sheet_name_length
        
        # Clean the name
        name = re.sub(r'[:\\/\?\*\[\]]', "", str(name)).strip()
        if not name:
            name = "Analysis"
        
        # Truncate if needed
        if len(name) > max_len:
            name = name[:max_len]
        
        # Ensure uniqueness
        final = name
        counter = 1
        while final in self.reserved_names:
            suffix = str(counter)
            final = name[:max_len - len(suffix)] + suffix
            counter += 1
            if counter > 99:  # Safety break
                break
        
        logger.debug(f"Generated sheet name: '{final}'")
        return final
    
    def generate_sheet_name(self, content_type: str, max_length: Optional[int] = None) -> str:
        """Legacy method - now calls clean_sheet_name for backward compatibility."""
        return self.clean_sheet_name(content_type)

    # =========================
    # Legacy compatibility methods
    # =========================

    def make_analysis_sheet(self, answer_text: str) -> Dict[str, Any]:
        """Legacy method for backward compatibility"""
        logger.info("Creating legacy analysis sheet")
        # FIX: Do not pass both `cells` (A1/A2) and `notes` (also writes A1/A2)
        # to _build_worksheet_payload — duplicate addresses cause COM error 0x8002802B
        # when the Office JS add-in tries to write conflicting values to the same cells.
        # Solution: use cells only, leave notes empty.
        cells = [
            {"address": "A1", "value": "AI Analysis"},
            {"address": "A2", "value": answer_text or ""},
            {"address": "A3", "value": "Analysis generated from the selected data preview."},
        ]
        return _build_worksheet_payload(
            name="AI_Analysis",
            cells=cells,
            formulas=[],
            notes=""  # FIX: empty — content moved into cells to avoid duplicate A1/A2
        )

    def create_fallback_structure(self, data: Dict[str, Any], question: str) -> Dict[str, Any]:
        """Fallback structure for error recovery"""
        logger.warning("Creating fallback structure due to errors")
        return {
            "sheets": [
                {
                    "name": "Analysis",
                    "type": "simple",
                    "cells": [
                        {"address": "A1", "value": "Simple Analysis"},
                        {"address": "A2", "value": question}
                    ],
                    "formulas": []
                }
            ],
            "metadata": {
                "created": datetime.now().isoformat(),
                "fallback": True,
                "question": question
            }
        }

    # =========================
    # Python-Only Results Sheet (no Excel formulas)
    # =========================
    
    def make_python_results_sheet(self, python_results: Dict[str, Any], hint: str = "") -> Dict[str, Any]:
        """
        Creates a sheet with pre-calculated Python results - NO Excel formulas needed.
        For top N/bottom N queries that are fully computed in Python.
        """
        print(f"[PYTHON-SHEET] Creating Python-only results sheet")
        print(f"[PYTHON-SHEET] Result type: {python_results.get('type', 'unknown')}")
        
        # Generate intelligent sheet name using LLM
        summary = python_results.get('summary', 'Python Analysis Results')
        sheet_name = self.generate_llm_sheet_name(hint, "Top N Analysis", {"header_analysis": {}}) if self.openai_client else "TopN_Results"
        
        cells = []
        
        # Professional title with better placement (A1)
        cells.append({
            "address": "A1",
            "value": summary,
            "style": {"bold": True, "fontSize": 16, "color": "#2E5984"}
        })
        
        # Subtitle explaining this is Python-computed (A2)
        cells.append({
            "address": "A2", 
            "value": "✓ Values pre-calculated in Python (no Excel formulas needed)",
            "style": {"italic": True, "fontSize": 10, "color": "#666666"}
        })
        
        # Results data (starting from row 4 for better spacing)
        data = python_results.get('data', [])
        if data:
            row = 4
            
            # Headers for data table with professional styling
            headers = list(data[0].keys()) if data else []
            for col_idx, header in enumerate(headers):
                cells.append({
                    "address": f"{chr(65 + col_idx)}{row}",
                    "value": header.replace('_', ' ').title(),
                    "style": {"bold": True, "backgroundColor": "#4A90E2", "color": "white", "fontSize": 12}
                })
            
            row += 1
            
            # Data rows with alternating row colors for better readability
            for row_idx, data_row in enumerate(data):
                bg_color = "#F8F9FA" if row_idx % 2 == 0 else "white"
                for col_idx, header in enumerate(headers):
                    value = data_row.get(header, "")
                    # Format numbers nicely
                    if isinstance(value, float):
                        value = round(value, 2)
                    
                    cells.append({
                        "address": f"{chr(65 + col_idx)}{row}",
                        "value": value,
                        "style": {"backgroundColor": bg_color}
                    })
                row += 1
        
        # Summary info at the bottom (don't show raw Python metadata)
        metadata = python_results.get('metadata', {})
        if metadata and 'original_rows' in metadata:
            bottom_row = row + 2
            cells.append({
                "address": f"A{bottom_row}",
                "value": f"Analysis Summary:",
                "style": {"bold": True, "fontSize": 11}
            })
            cells.append({
                "address": f"A{bottom_row + 1}",
                "value": f"• Analyzed {metadata.get('original_rows', 0)} total records",
                "style": {"fontSize": 10}
            })
            if 'filtered_rows' in metadata and metadata['filtered_rows'] != metadata.get('original_rows'):
                cells.append({
                    "address": f"A{bottom_row + 2}",
                    "value": f"• Filtered to {metadata.get('filtered_rows', 0)} relevant records",
                    "style": {"fontSize": 10}
                })
        
        print(f"[PYTHON-SHEET] Created {len(cells)} cells with professional formatting")
        
        return {
            "name": sheet_name,
            "cells": cells,
            "formulas": [],  # NO Excel formulas - all values pre-calculated in Python
            "notes": f"Python-computed analysis for: {hint}. All calculations performed in Python for accuracy."
        }

    def _convert_indices_to_names(self, col_list: List[str], available_columns: List[str]) -> List[str]:
        """Convert column indices back to column names if they're detected."""
        converted = []
        for col in col_list:
            if isinstance(col, str) and col.isdigit():
                # This is likely a column index
                idx = int(col)
                if 0 <= idx < len(available_columns):
                    converted.append(available_columns[idx])
                else:
                    converted.append(col)
            else:
                # This is already a column name
                converted.append(col)
        return converted

    # =========================
    # Calculated Values Sheet (from previous conversation)
    # =========================

    def make_calculated_values_sheet(self, data: Dict[str, Any], hint: str = "", source_sheet: str = "Sheet1") -> Dict[str, Any]:
        """
        Creates a sheet with pre-calculated values and human-readable descriptions.
        """
        #logger.info(f"Creating calculated values sheet for query: {hint}")
        
        header_map = self.map_headers_to_letters(data)
        ha = data.get("header_analysis", {}) or {}
        ps = data.get("pattern_sample", [])
        
        # Check if we have LLM-derived multi-aggregation pattern (highest priority)
        llm_multi_agg_pattern = data.get("llm_multi_aggregation_pattern")
        if (llm_multi_agg_pattern and 
            llm_multi_agg_pattern.get("aggregation_details") and 
            llm_multi_agg_pattern.get("groupby_columns")):
            logger.debug("Using LLM-derived multi-aggregation pattern")
            aggregation_details = llm_multi_agg_pattern["aggregation_details"]
            groupby_columns = llm_multi_agg_pattern["groupby_columns"]
            
            # Extract all metric columns from aggregation details
            numeric_cols = [agg.get("column") for agg in aggregation_details if agg.get("column")]
            categorical_cols = groupby_columns
            
            # Create a multi-aggregation groupby pattern
            groupby_pattern = {
                "metric_column": numeric_cols[0],  # Primary for backward compatibility
                "groupby_column": groupby_columns[0],
                "aggregation": aggregation_details[0].get("function", "SUM").upper(),
                "multi_aggregation": True,
                "aggregation_details": aggregation_details,
                "groupby_columns": groupby_columns
            }
            selection_reasoning = f"Using LLM multi-aggregation plan with {len(aggregation_details)} aggregations"
        
        # Check if we have LLM-derived single groupby pattern (fallback)
        elif data.get("llm_groupby_pattern"):
            llm_groupby_pattern = data.get("llm_groupby_pattern")
            if (llm_groupby_pattern.get("metric_column") is not None and 
                llm_groupby_pattern.get("groupby_column") is not None):
                logger.debug("Using LLM-derived single groupby pattern")
                groupby_pattern = llm_groupby_pattern
                # Extract columns from the pattern for compatibility
                numeric_cols = [llm_groupby_pattern["metric_column"]]
                categorical_cols = [llm_groupby_pattern["groupby_column"]]
                selection_reasoning = "Using LLM analysis plan for reliable column selection"
        else:
            # Use LLM-based intelligent column selection (fallback)
            column_selection = select_relevant_columns(hint, ha)
            logger.debug("Running heuristic column selection")
            
            # SAFEGUARD: Convert column indices back to names if they're detected
            available_columns = list(ha.keys())
            numeric_cols = self._convert_indices_to_names(column_selection["numeric_cols"], available_columns)
            categorical_cols = self._convert_indices_to_names(column_selection["categorical_cols"], available_columns)
            
            selection_reasoning = column_selection["reasoning"]
            groupby_pattern = column_selection.get("groupby_pattern")
            
            # Fix groupby_pattern if it contains indices
            if groupby_pattern:
                if groupby_pattern.get("metric_column") and groupby_pattern["metric_column"].isdigit():
                    idx = int(groupby_pattern["metric_column"])
                    if 0 <= idx < len(available_columns):
                        groupby_pattern["metric_column"] = available_columns[idx]
                
                if groupby_pattern.get("groupby_column") and groupby_pattern["groupby_column"].isdigit():
                    idx = int(groupby_pattern["groupby_column"])
                    if 0 <= idx < len(available_columns):
                        groupby_pattern["groupby_column"] = available_columns[idx]
            
            logger.debug(f"Column analysis complete: {len(numeric_cols)} numeric, {len(categorical_cols)} categorical")
        
        logger.debug(f"LLM selected {len(numeric_cols)} numeric and {len(categorical_cols)} categorical columns")
        logger.debug(f"Selection reasoning: {selection_reasoning}")
        
        if groupby_pattern:
            logger.info(f"Detected groupby pattern: {groupby_pattern['aggregation']} {groupby_pattern['metric_column']} by {groupby_pattern['groupby_column']}")
        
        cells = []
        formulas = []
        row_counter = 1
        
        # Title and column selection info
        if groupby_pattern and groupby_pattern.get("multi_aggregation"):
            # Multi-aggregation: Will have custom headers in row 3
            cells.extend([
                {"address": "A1", "value": "📊 Multi-Aggregation Analytics"},
                {"address": "A2", "value": f"🎯 Analysis: {selection_reasoning}"}
            ])
        else:
            # Single aggregation: Traditional layout
            cells.extend([
                {"address": "A1", "value": "📊 Analytics Summary"},
                {"address": "B1", "value": "Excel Formula"},
                {"address": "A2", "value": f"🎯 Analysis: {selection_reasoning}"}
            ])
        row_counter = 4
        print(f"In Make_calulated_sheet after cells extend")
        # Process groupby patterns or individual metrics
        try:
            if groupby_pattern:
                # Handle aggregation patterns like "sum sales by sku"
                calculated_results = self._calculate_grouped_aggregation(ps, groupby_pattern, source_sheet, header_map)
                print(f"Calculated {len(calculated_results)} grouped aggregation results")
            else:
                # Original behavior for individual column metrics
                calculated_results = self._calculate_values_from_sample(ps, numeric_cols, categorical_cols, source_sheet)
                print(f"Calculated {len(calculated_results)} individual column metrics from sample data")
        except Exception as e:
            logger.error(f"Failed to calculate values from sample: {str(e)}")
            calculated_results = []
        print(f"In Make_calulated_sheet after calculated_results")
        # Recent fix: Handle multi-aggregation columnar layout vs single aggregation
        if groupby_pattern and groupby_pattern.get("multi_aggregation"):
            # Multi-aggregation: Create proper columnar layout
            self._populate_multi_aggregation_columnar(calculated_results, cells, formulas, groupby_pattern)
        else:
            # Single aggregation: Use existing row-based layout
            for metric in calculated_results:
                # Human readable description
                cells.append({
                    "address": f"A{row_counter}", 
                    "value": metric["description"]
                })
                
                # Recent fix: Use Excel formula as the value instead of Python calculation
                formulas.append({
                    "address": f"B{row_counter}",
                    "formula": metric["formula"],
                    "description": f"Formula for {metric['description']}"
                })
                
                row_counter += 1
        
        # Convert name ranges to letters if available
        if header_map:
            for f in formulas:
                f["formula"] = self.convert_name_range_to_letter(f.get("formula", ""), header_map)
        
        # Recent fix: Correct formula row references to match actual placement
        for f in formulas:
            # Extract row number from formula address (e.g., "C4" -> 4)
            address = f.get("address", "")
            if address and len(address) > 1:
                try:
                    target_row = int(address[1:])  # Get row number from address like "C4"
                    f["formula"] = self._correct_formula_row_references(f.get("formula", ""), target_row)
                except (ValueError, IndexError):
                    # If we can't extract row number, skip correction
                    pass
        
        # Generate intelligent sheet name
        sheet_name = self.generate_llm_sheet_name(hint, "Formula Analysis", data) if self.openai_client else "Formula_Sheet"
        
        result = _build_worksheet_payload(
            name=sheet_name,
            cells=cells,
            formulas=formulas,
            notes=f"Analysis for: {hint}. Calculated values with Excel formulas for verification."
        )
        
        logger.info(f"Calculated values sheet created with {len(calculated_results)} metrics")
        return result

    def _calculate_values_from_sample(self, pattern_sample: List[List[Any]], numeric_cols: List[str], categorical_cols: List[str], source_sheet: str = "Sheet1") -> List[Dict[str, Any]]:
        """Calculate actual values from the pattern_sample data."""
        logger.info("Calculating values from pattern sample")
        
        logger.debug("Starting value calculation from sample data")

        
        if not pattern_sample or len(pattern_sample) < 2:
            logger.warning("Insufficient sample data for calculations")

            return []
        
        headers = pattern_sample[0] if isinstance(pattern_sample[0], list) else []
        data_rows = pattern_sample[1:] if len(pattern_sample) > 1 else []
        
        if not headers or not data_rows:
            logger.warning("No headers or data rows found in sample")
            return []
        
        results = []
        header_idx = {header: idx for idx, header in enumerate(headers)}
        
        
        # Calculate metrics for numeric columns
        for col_name in numeric_cols:
            # Recent fix: Convert column index to name if needed
            if isinstance(col_name, str) and col_name.isdigit():
                idx = int(col_name)
                if 0 <= idx < len(headers):
                    col_name = headers[idx]
                else:
                    continue
            
            
            if col_name not in header_idx:
                logger.debug(f"Skipping numeric column '{col_name}' - not found in headers")
                continue
            
            col_idx = header_idx[col_name]
            values = []
            
            # Extract numeric values
            for row_idx, row in enumerate(data_rows):
                if col_idx < len(row) and row[col_idx] is not None:
                    try:
                        val = float(row[col_idx]) if row[col_idx] != '' else 0
                        values.append(val)
                    except (ValueError, TypeError):
                        continue
            
            
            if values:
                logger.debug(f"Processing {len(values)} values for column '{col_name}'")
                
                # Calculate various metrics
                
                metrics = [
                    ("Sum", sum(values), "SUM"),
                    ("Average", sum(values) / len(values), "AVERAGE"),
                    ("Maximum", max(values), "MAX"),
                    ("Minimum", min(values), "MIN"),
                    ("Count", len(values), "COUNT")
                ]
                
                
                for metric_name, value, formula_type in metrics:
                    result_entry = {
                        "description": f"{metric_name} of {col_name}",
                        "value": round(value, 2) if isinstance(value, float) else value,
                        "formula": f"={formula_type}({source_sheet}!{col_name}:{col_name})"
                    }
                    results.append(result_entry)
            else:
                logger.debug(f"No valid values found for numeric column {col_name}")
        
        # Calculate metrics for categorical columns
        for col_name in categorical_cols:
            # Recent fix: Convert column index to name if needed
            if isinstance(col_name, str) and col_name.isdigit():
                idx = int(col_name)
                if 0 <= idx < len(headers):
                    col_name = headers[idx]
                else:
                    continue
            
            
            if col_name not in header_idx:
                logger.debug(f"Skipping categorical column '{col_name}' - not found in headers")
                continue
                
            col_idx = header_idx[col_name]
            values = []
            
            # Extract non-empty values
            for row_idx, row in enumerate(data_rows):
                if col_idx < len(row) and row[col_idx] is not None and row[col_idx] != '':
                    values.append(str(row[col_idx]))
            
            
            if values:
                logger.debug(f"Processing {len(values)} values for categorical column '{col_name}'")
                
                
                # Unique count
                unique_values = set(values)
                unique_count = len(unique_values)
                
                unique_result = {
                    "description": f"Unique {col_name} categories",
                    "value": unique_count,
                    "formula": f"=COUNTA(UNIQUE({source_sheet}!{col_name}:{col_name}))-2"
                }
                results.append(unique_result)
                
                # Total count
                total_result = {
                    "description": f"Total {col_name} entries",
                    "value": len(values),
                    "formula": f"=COUNTA({source_sheet}!{col_name}:{col_name})"
                }
                results.append(total_result)
                
                # Most frequent value
                counter = Counter(values)
                most_common = counter.most_common(1)
                if most_common:
                    most_frequent_value, frequency = most_common[0]
                    
                    frequent_result = {
                        "description": f"Most frequent {col_name}",
                        "value": f"{most_frequent_value} ({frequency} times)",
                        "formula": f"=INDEX({source_sheet}!{col_name}:{col_name},MODE(IF({source_sheet}!{col_name}:{col_name}<>\"\",MATCH({source_sheet}!{col_name}:{col_name},{source_sheet}!{col_name}:{col_name},0))))"
                    }
                    results.append(frequent_result)
            else:
                logger.debug(f"No valid values found for categorical column {col_name}")
        
        logger.debug(f"Calculated {len(results)} total metrics")
        
        logger.debug("Value calculation completed")

        
        return results
    
    def _calculate_grouped_aggregation(self, pattern_sample: List[List[Any]], groupby_pattern: Dict[str, Any], source_sheet: str = "Sheet1", header_map: Dict[str, str] = None) -> List[Dict[str, Any]]:
        """
        Calculate grouped aggregations like 'sum sales by sku' from pattern_sample data.
        
        Args:
            pattern_sample: List of lists with headers as first row, data as subsequent rows
            groupby_pattern: Dict with keys: metric_column, groupby_column, aggregation
            source_sheet: Source sheet name for formula generation
        """
        
        if not pattern_sample or len(pattern_sample) < 2:
            return []
        
        headers = pattern_sample[0] if isinstance(pattern_sample[0], list) else []
        data_rows = pattern_sample[1:] if len(pattern_sample) > 1 else []
        
        # Recent fix: Handle both single and multi-aggregation patterns
        is_multi_agg = groupby_pattern.get("multi_aggregation", False)
        
        if is_multi_agg and groupby_pattern.get("aggregation_details"):
            # Multi-aggregation case: adapt format and use existing multi-aggregation method
            adapted_pattern = {
                "aggregations": [
                    {"column": agg.get("column"), "function": agg.get("function", "SUM")}
                    for agg in groupby_pattern["aggregation_details"]
                ],
                "groupby_columns": groupby_pattern.get("groupby_columns", []),
                "filters": []
            }
            logger.info(f"Using multi-aggregation with {len(adapted_pattern['aggregations'])} aggregations")
            multi_agg_result = self._calculate_multi_aggregation_grouped(pattern_sample, adapted_pattern, source_sheet, header_map)
            
            # Recent fix: Create columnar format for multi-aggregation results
            return self._format_multi_aggregation_columnar(multi_agg_result, groupby_pattern)
        
        # Single aggregation case (original behavior)
        metric_col = groupby_pattern["metric_column"]
        group_col = groupby_pattern["groupby_column"]
        aggregation = groupby_pattern["aggregation"]
        
        # Handle case where column references are indices instead of names
        if isinstance(metric_col, str) and metric_col.isdigit():
            idx = int(metric_col)
            if 0 <= idx < len(headers):
                metric_col = headers[idx]
            else:
                logger.warning(f"Metric column index {idx} out of range for {len(headers)} headers")
                
        if isinstance(group_col, str) and group_col.isdigit():
            idx = int(group_col)
            if 0 <= idx < len(headers):
                group_col = headers[idx]
            else:
                logger.warning(f"Group column index {idx} out of range for {len(headers)} headers")
        
        if not headers or not data_rows:
            return []
        
        header_idx = {header: idx for idx, header in enumerate(headers)}
        
        if metric_col not in header_idx or group_col not in header_idx:
            return []
        
                # Recent fix: Convert column names to letters for Excel formulas
        metric_col_letter = metric_col
        group_col_letter = group_col




        if header_map:
            if metric_col in header_map:
                metric_col_letter = header_map[metric_col]

            else:
                logger.debug(f"metric_col '{metric_col}' not found in header_map")

            if group_col in header_map:
                group_col_letter = header_map[group_col]

            else:
                logger.debug(f"group_col '{group_col}' not found in header_map")

        else:
            pass

        metric_idx = header_idx[metric_col]
        group_idx = header_idx[group_col]
        
        # Group data by the groupby column
        groups = {}
        for row in data_rows:
            if group_idx < len(row) and metric_idx < len(row):
                group_value = row[group_idx]
                metric_value = row[metric_idx]
                
                # Skip empty or null group values
                if group_value is None or group_value == '':
                    continue
                
                group_key = str(group_value)
                
                # For COUNT operations, we don't need numeric conversion of metric values
                if aggregation in ["COUNT", "FREQUENCY"]:
                    # For COUNT, just track non-empty metric values
                    if metric_value is not None and metric_value != '':
                        if group_key not in groups:
                            groups[group_key] = []
                        groups[group_key].append(1)  # Just count occurrences
                else:
                    # For other aggregations, we need numeric values
                    if metric_value is None or metric_value == '':
                        continue
                    try:
                        metric_num = float(metric_value)
                        if group_key not in groups:
                            groups[group_key] = []
                        groups[group_key].append(metric_num)
                    except (ValueError, TypeError):
                        continue
        
        
        # Calculate aggregation for each group
        results = []
        result_row = 4  # Starting row for results (after headers)
        for group_name, values in groups.items():
            if not values:
                continue
            
            # Convert to numeric values and handle statistics
            import statistics
            import math
            
            # Basic Aggregations
            if aggregation in ["SUM", "TOTAL"]:
                agg_value = sum(values)
            elif aggregation in ["AVERAGE", "MEAN", "AVG"]:
                agg_value = sum(values) / len(values)
            elif aggregation in ["MAX", "MAXIMUM"]:
                agg_value = max(values)
            elif aggregation in ["MIN", "MINIMUM"]:
                agg_value = min(values)
            elif aggregation in ["COUNT", "FREQUENCY"]:
                agg_value = len(values)
                
            # Statistical Aggregations
            elif aggregation == "MEDIAN":
                agg_value = statistics.median(values)
            elif aggregation in ["STDEV", "STD", "STANDARD_DEVIATION"]:
                agg_value = statistics.stdev(values) if len(values) > 1 else 0
            elif aggregation in ["VAR", "VARIANCE"]:
                agg_value = statistics.variance(values) if len(values) > 1 else 0
            elif aggregation == "MODE":
                try:
                    agg_value = statistics.mode(values)
                except statistics.StatisticsError:
                    agg_value = values[0]  # No unique mode
                    
            # Percentiles and Quantiles
            elif aggregation in ["Q1", "FIRST_QUARTILE"]:
                sorted_vals = sorted(values)
                agg_value = statistics.quantiles(sorted_vals, n=4)[0] if len(values) > 1 else values[0]
            elif aggregation in ["Q3", "THIRD_QUARTILE"]:
                sorted_vals = sorted(values)
                agg_value = statistics.quantiles(sorted_vals, n=4)[2] if len(values) > 1 else values[0]
            elif aggregation in ["IQR", "INTERQUARTILE_RANGE"]:
                if len(values) > 1:
                    sorted_vals = sorted(values)
                    q1, q3 = statistics.quantiles(sorted_vals, n=4)[0], statistics.quantiles(sorted_vals, n=4)[2]
                    agg_value = q3 - q1
                else:
                    agg_value = 0
            elif aggregation in ["P90", "90TH_PERCENTILE"]:
                sorted_vals = sorted(values)
                idx = int(0.9 * (len(values) - 1))
                agg_value = sorted_vals[idx]
            elif aggregation in ["P95", "95TH_PERCENTILE"]:
                sorted_vals = sorted(values)
                idx = int(0.95 * (len(values) - 1))
                agg_value = sorted_vals[idx]
            elif aggregation in ["P99", "99TH_PERCENTILE"]:
                sorted_vals = sorted(values)
                idx = int(0.99 * (len(values) - 1))
                agg_value = sorted_vals[idx]
                
            # Range and Spread
            elif aggregation == "RANGE":
                agg_value = max(values) - min(values)
            elif aggregation in ["UNIQUE_COUNT", "DISTINCT_COUNT"]:
                agg_value = len(set(values))
                
            # Business Analytics
            elif aggregation in ["CV", "COEFFICIENT_OF_VARIATION"]:
                if len(values) > 1:
                    mean_val = sum(values) / len(values)
                    std_val = statistics.stdev(values)
                    agg_value = std_val / mean_val if mean_val != 0 else 0
                else:
                    agg_value = 0
            elif aggregation in ["SKEW", "SKEWNESS"]:
                # Simplified skewness calculation
                if len(values) > 2:
                    mean_val = sum(values) / len(values)
                    std_val = statistics.stdev(values)
                    if std_val != 0:
                        skew = sum(((x - mean_val) / std_val) ** 3 for x in values) / len(values)
                        agg_value = skew
                    else:
                        agg_value = 0
                else:
                    agg_value = 0
                    
            # Growth and Change
            elif aggregation in ["FIRST", "EARLIEST"]:
                agg_value = values[0]
            elif aggregation in ["LAST", "LATEST"]:
                agg_value = values[-1]
                
            # Custom Counts
            elif aggregation == "POSITIVE_COUNT":
                agg_value = sum(1 for x in values if x > 0)
            elif aggregation == "NEGATIVE_COUNT":
                agg_value = sum(1 for x in values if x < 0)
            elif aggregation == "ZERO_COUNT":
                agg_value = sum(1 for x in values if x == 0)
            elif aggregation == "NON_ZERO_COUNT":
                agg_value = sum(1 for x in values if x != 0)
                
            else:
                agg_value = sum(values)  # Default to SUM
            
            # Create result entry with appropriate Excel formula based on aggregation type
            # Use correct column references for dynamic formulas
            # Recent fix: The criteria should reference column A (where group names are stored) in the same row as the formula
            # Note: The actual row will be determined when the formula is placed in the worksheet
            criteria_cell = f"A{result_row}"  # Reference the cell containing the group name in column A

            
            if aggregation in ["SUM", "TOTAL"]:
                formula = f"=SUMIF({source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter})"


            elif aggregation in ["AVERAGE", "MEAN", "AVG"]:
                formula = f"=AVERAGEIF({source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter})"
            elif aggregation in ["COUNT", "FREQUENCY"]:
                formula = f"=COUNTIF({source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell})"
            elif aggregation in ["MAX", "MAXIMUM"]:
                formula = f"=MAXIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell})"
            elif aggregation in ["MIN", "MINIMUM"]:
                formula = f"=MINIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell})"
                
            # Statistical Functions (require array formulas or complex expressions)
            elif aggregation == "MEDIAN":
                formula = f"=MEDIAN(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["STDEV", "STD", "STANDARD_DEVIATION"]:
                formula = f"=STDEV.P(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["VAR", "VARIANCE"]:
                formula = f"=VAR.P(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation == "MODE":
                formula = f"=MODE.SNGL(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
                
            # Percentiles and Quantiles
            elif aggregation in ["Q1", "FIRST_QUARTILE"]:
                formula = f"=QUARTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),1)"
            elif aggregation in ["Q3", "THIRD_QUARTILE"]:
                formula = f"=QUARTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),3)"
            elif aggregation in ["IQR", "INTERQUARTILE_RANGE"]:
                formula = f"=QUARTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),3)-QUARTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),1)"
            elif aggregation in ["P90", "90TH_PERCENTILE"]:
                formula = f"=PERCENTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),0.9)"
            elif aggregation in ["P95", "95TH_PERCENTILE"]:
                formula = f"=PERCENTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),0.95)"
            elif aggregation in ["P99", "99TH_PERCENTILE"]:
                formula = f"=PERCENTILE.INC(IF({source_sheet}!{group_col_letter}:{group_col_letter}={criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter}),0.99)"
                
            # Range and Spread
            elif aggregation == "RANGE":
                formula = f"=MAXIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{group_col_letter}:{group_col_letter},\"{group_name}\")-MINIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{group_col_letter}:{group_col_letter},\"{group_name}\")"
            elif aggregation in ["UNIQUE_COUNT", "DISTINCT_COUNT"]:
                formula = f"=SUMPRODUCT((IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",1/COUNTIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{metric_col_letter}:{metric_col_letter},{source_sheet}!{group_col_letter}:{group_col_letter},\"{group_name}\"),0)))"
                
            # Business Analytics
            elif aggregation in ["CV", "COEFFICIENT_OF_VARIATION"]:
                formula = f"=STDEV.P(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}))/AVERAGE(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["SKEW", "SKEWNESS"]:
                formula = f"=SKEW(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["KURT", "KURTOSIS"]:
                formula = f"=KURT(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
                
            # Growth and Change
            elif aggregation in ["FIRST", "EARLIEST"]:
                formula = f"=INDEX(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}),1)"
            elif aggregation in ["LAST", "LATEST"]:
                formula = f"=INDEX(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter}),COUNTA(IF({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\",{source_sheet}!{metric_col_letter}:{metric_col_letter})))"
                
            # Custom Counts
            elif aggregation == "POSITIVE_COUNT":
                formula = f"=SUMPRODUCT(({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\")*({source_sheet}!{metric_col_letter}:{metric_col_letter}>0))"
            elif aggregation == "NEGATIVE_COUNT":
                formula = f"=SUMPRODUCT(({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\")*({source_sheet}!{metric_col_letter}:{metric_col_letter}<0))"
            elif aggregation == "ZERO_COUNT":
                formula = f"=SUMPRODUCT(({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\")*({source_sheet}!{metric_col_letter}:{metric_col_letter}=0))"
            elif aggregation == "NON_ZERO_COUNT":
                formula = f"=SUMPRODUCT(({source_sheet}!{group_col_letter}:{group_col_letter}=\"{group_name}\")*({source_sheet}!{metric_col_letter}:{metric_col_letter}<>0))"
                
            else:
                # Fallback to SUM for unknown aggregations
                formula = f"=SUMIF({source_sheet}!{group_col_letter}:{group_col_letter},{criteria_cell},{source_sheet}!{metric_col_letter}:{metric_col_letter})"
            
            result_entry = {
                "description": f"{group_name}",
                "value": round(agg_value, 2) if isinstance(agg_value, float) else agg_value,
                "formula": formula
            }
            
            results.append(result_entry)
            #print(f"[LOG] Added grouped result: {group_name} = {agg_value}")
            result_row += 1  # Increment for next result
        
        # Sort results by value in descending order for better presentation
        results.sort(key=lambda x: x["value"], reverse=True)
        
        return results
    
    def _format_multi_aggregation_columnar(self, multi_agg_result: Dict[str, Any], groupby_pattern: Dict[str, Any]) -> Dict[str, Any]:
        """
        Format multi-aggregation results for proper columnar layout.
        
        Recent fix: Returns structured data for columnar display instead of row-based format.
        
        Args:
            multi_agg_result: Result from _calculate_multi_aggregation_grouped
            groupby_pattern: Original groupby pattern with aggregation details
            
        Returns:
            Dictionary with structured columnar data
        """
        aggregation_details = groupby_pattern.get("aggregation_details", [])
        groupby_columns = groupby_pattern.get("groupby_columns", [])
        
        if not multi_agg_result.get("results") or not aggregation_details:
            return {"groups": {}, "aggregations": [], "formulas": {}}
        
        # Extract group keys and organize formulas by group and aggregation
        groups = {}
        formulas_by_group = {}
        
        for result in multi_agg_result.get("results", []):
            group_key = result.get("group_key", "Unknown")
            
            agg_results = result.get("aggregation_results", {})
            groups[group_key] = {}
            formulas_by_group[group_key] = {}
            
            for agg_detail in aggregation_details:
                function = agg_detail.get("function", "SUM").upper()
                column = agg_detail.get("column", "Unknown")
                agg_key = f"{column}_{function}"
                agg_desc = f"{function} of {column}"
                
                if agg_key in agg_results:
                    formula = agg_results[agg_key].get("formula", "")
                    groups[group_key][agg_desc] = formula
                    formulas_by_group[group_key][agg_key] = formula
        
        return {
            "groups": groups,
            "aggregations": aggregation_details,
            "formulas": formulas_by_group,
            "groupby_column": groupby_columns[0] if groupby_columns else "Group"
        }
    
    def _populate_multi_aggregation_columnar(self, calculated_results: List[Dict[str, Any]], cells: List[Dict[str, Any]], formulas: List[Dict[str, Any]], groupby_pattern: Dict[str, Any]):
        """
        Populate cells and formulas for multi-aggregation in proper columnar format.
        
        Recent fix: Creates the exact structure requested - one row per group with columns for each aggregation.
        
        Args:
            calculated_results: Results from multi-aggregation processing
            cells: List to populate with cell data
            formulas: List to populate with formula data
            groupby_pattern: Pattern containing aggregation details
        """
        if not calculated_results:
            return
        
        # Extract the structured data
        # Recent fix: calculated_results is now a dictionary from _format_multi_aggregation_columnar
        if isinstance(calculated_results, dict):
            # New dictionary format
            groups = calculated_results.get("groups", {})
            aggregation_details = groupby_pattern.get("aggregation_details", [])
        elif hasattr(calculated_results, '__iter__') and len(calculated_results) > 0 and isinstance(calculated_results[0], dict):
            # Old list format - fallback for compatibility
            groups = {}
            aggregation_details = groupby_pattern.get("aggregation_details", [])
            
            for result in calculated_results:
                desc = result.get("description", "")
                formula = result.get("formula", "")
                
                # Parse description to extract group and aggregation
                if " - " in desc and " of " in desc:
                    parts = desc.split(" - ")
                    if len(parts) >= 2:
                        group_key = parts[0]
                        agg_part = parts[1]
                        
                        if group_key not in groups:
                            groups[group_key] = {}
                        
                        groups[group_key][agg_part] = formula
        else:
            # Unexpected format
            print(f"[ERROR] Unexpected calculated_results format: {type(calculated_results)}")
            return
        
        # Create header row - Recent fix: Include all groupby columns as separate headers
        groupby_columns = groupby_pattern.get("groupby_columns", ["Group"])
        header_row = list(groupby_columns)  # All groupby columns as separate headers
        
        for agg_detail in aggregation_details:
            function = agg_detail.get("function", "SUM").upper()
            column = agg_detail.get("column", "Unknown")
            header_row.append(f"{function}({column})")
        
        # Add headers as cells
        for col_idx, header in enumerate(header_row):
            cells.append({
                "address": f"{chr(65 + col_idx)}3",  # A3, B3, C3, etc.
                "value": header
            })
        
        # Add data rows
        row_num = 4  # Start from row 4
        for group_key, group_formulas in groups.items():
            # Recent fix: Parse group_key to extract individual column values
            # group_key format: "sku_code:ROMDT138ZL49-664A-30 | warehouse_id:WH01"
            group_parts = group_key.split(" | ")
            group_values = {}
            
            for part in group_parts:
                if ":" in part:
                    col_name, col_value = part.split(":", 1)
                    group_values[col_name] = col_value
            
            # Place each groupby column value in its respective column
            col_idx = 0
            for groupby_col in groupby_columns:
                value = group_values.get(groupby_col, "Unknown")
                cells.append({
                    "address": f"{chr(65 + col_idx)}{row_num}",
                    "value": value
                })
                col_idx += 1
            
            # Formulas in subsequent columns (after all groupby columns)
            for agg_detail in aggregation_details:
                function = agg_detail.get("function", "SUM").upper()
                column = agg_detail.get("column", "Unknown")
                agg_desc = f"{function} of {column}"
                
                if agg_desc in group_formulas:
                    formulas.append({
                        "address": f"{chr(65 + col_idx)}{row_num}",  # B4, C4, etc.
                        "formula": group_formulas[agg_desc],
                        "description": f"{group_key} - {agg_desc}"
                    })
                col_idx += 1
            
            row_num += 1
        
        logger.info(f"Created columnar layout with {len(groups)} groups and {len(aggregation_details)} aggregations")
    
    def _calculate_multi_aggregation_grouped(self, pattern_sample: List[List[Any]], 
                                           multi_groupby_pattern: Dict[str, Any], 
                                           source_sheet: str = "Sheet1", 
                                           header_map: Dict[str, str] = None) -> Dict[str, Any]:
        """
        Calculate multiple aggregations simultaneously across multiple grouping columns.
        
        Example: "sum qty and average revenue by sku_code and id"
        
        Args:
            pattern_sample: List of lists with headers as first row, data as subsequent rows
            multi_groupby_pattern: Dict with structure:
                {
                    "aggregations": [
                        {"column": "qty", "function": "SUM"},
                        {"column": "revenue", "function": "AVERAGE"}
                    ],
                    "groupby_columns": ["sku_code", "id"],
                    "filters": [
                        {"column": "region", "operator": "=", "value": "North"}
                    ] (optional)
                }
            source_sheet: Source sheet name for formula generation
            header_map: Mapping of column names to Excel letters
            
        Returns:
            {
                "results": [
                    {
                        "group_key": "group_identifier",
                        "group_values": {"sku_code": "SKU001", "id": "A001"},
                        "aggregation_results": {
                            "qty_SUM": {"value": 100, "formula": "=SUMIFS(...)"},
                            "revenue_AVERAGE": {"value": 1500.0, "formula": "=AVERAGEIFS(...)"}
                        }
                    }
                ],
                "summary": {
                    "total_groups": 10,
                    "aggregations_count": 2,
                    "groupby_columns": ["sku_code", "id"]
                }
            }
        """
        
        if not pattern_sample or len(pattern_sample) < 2:
            return {"results": [], "summary": {"total_groups": 0, "aggregations_count": 0, "groupby_columns": []}}
        
        headers = pattern_sample[0] if isinstance(pattern_sample[0], list) else []
        data_rows = pattern_sample[1:] if len(pattern_sample) > 1 else []
        
        # Extract configuration
        aggregations = multi_groupby_pattern.get("aggregations", [])
        groupby_columns = multi_groupby_pattern.get("groupby_columns", [])
        filters = multi_groupby_pattern.get("filters", [])
        
        if not aggregations or not groupby_columns:
            return {"results": [], "summary": {"total_groups": 0, "aggregations_count": 0, "groupby_columns": []}}
        
        
        # Validate columns exist
        header_idx = {header: idx for idx, header in enumerate(headers)}
        
        # Validate aggregation columns
        valid_aggregations = []
        for agg in aggregations:
            col = agg.get("column")
            func = agg.get("function", "SUM").upper()
            
            # Handle column index conversion
            if isinstance(col, str) and col.isdigit():
                idx = int(col)
                if 0 <= idx < len(headers):
                    col = headers[idx]
            
            if col in header_idx:
                valid_aggregations.append({"column": col, "function": func})
            else:
                logger.warning(f"Aggregation column '{col}' not found in headers")
        
        # Validate groupby columns  
        valid_groupby_columns = []
        for col in groupby_columns:
            # Handle column index conversion
            if isinstance(col, str) and col.isdigit():
                idx = int(col)
                if 0 <= idx < len(headers):
                    col = headers[idx]
            
            if col in header_idx:
                valid_groupby_columns.append(col)
            else:
                logger.warning(f"Groupby column '{col}' not found in headers")
        
        if not valid_aggregations or not valid_groupby_columns:
            return {"results": [], "summary": {"total_groups": 0, "aggregations_count": 0, "groupby_columns": []}}
        
        # Apply filters to data if specified
        filtered_data_rows = self._apply_filters_to_data(data_rows, headers, filters) if filters else data_rows
        
        # Group data by multiple columns
        groups = {}
        for row in filtered_data_rows:
            if len(row) < len(headers):
                continue  # Skip incomplete rows
            
            # Create composite group key
            group_values = {}
            group_key_parts = []
            
            for group_col in valid_groupby_columns:
                group_idx = header_idx[group_col]
                if group_idx < len(row):
                    group_value = row[group_idx]
                    if group_value is None or group_value == '':
                        break  # Skip rows with missing group values
                    group_values[group_col] = group_value
                    group_key_parts.append(f"{group_col}:{group_value}")
                else:
                    break
            
            if len(group_key_parts) == len(valid_groupby_columns):
                group_key = " | ".join(group_key_parts)
                
                if group_key not in groups:
                    groups[group_key] = {
                        "group_values": group_values,
                        "rows": []
                    }
                groups[group_key]["rows"].append(row)
        
        
        # Calculate aggregations for each group
        results = []
        result_row = 4  # Starting row for Excel formulas
        
        for group_key, group_data in groups.items():
            group_values = group_data["group_values"]
            rows = group_data["rows"]
            
            aggregation_results = {}
            
            for agg in valid_aggregations:
                agg_col = agg["column"]
                agg_func = agg["function"]
                agg_idx = header_idx[agg_col]
                
                # Extract values for this aggregation
                values = []
                for row in rows:
                    if agg_idx < len(row):
                        value = row[agg_idx]
                        
                        if agg_func in ["COUNT", "FREQUENCY"]:
                            # For COUNT, just check non-empty values
                            if value is not None and value != '':
                                values.append(1)
                        else:
                            # For other functions, need numeric values
                            if value is not None and value != '':
                                try:
                                    numeric_value = float(value)
                                    values.append(numeric_value)
                                except (ValueError, TypeError):
                                    continue
                
                if not values:
                    continue
                
                # Calculate aggregation
                calculated_value = self._calculate_single_aggregation(values, agg_func)
                
                # Generate Excel formula
                formula = self._generate_multi_criteria_formula(
                    agg_func, agg_col, valid_groupby_columns, group_values, 
                    source_sheet, header_map, result_row
                )
                
                # Store result
                agg_key = f"{agg_col}_{agg_func}"
                aggregation_results[agg_key] = {
                    "value": round(calculated_value, 2) if isinstance(calculated_value, float) else calculated_value,
                    "formula": formula,
                    "column": agg_col,
                    "function": agg_func
                }
            
            if aggregation_results:
                results.append({
                    "group_key": group_key,
                    "group_values": group_values,
                    "aggregation_results": aggregation_results,
                    "row_count": len(rows)
                })
                result_row += 1
        
        # Sort results by first aggregation value for consistent presentation
        if results and valid_aggregations:
            first_agg_key = f"{valid_aggregations[0]['column']}_{valid_aggregations[0]['function']}"
            results.sort(
                key=lambda x: x["aggregation_results"].get(first_agg_key, {}).get("value", 0), 
                reverse=True
            )
        
        summary = {
            "total_groups": len(results),
            "aggregations_count": len(valid_aggregations),
            "groupby_columns": valid_groupby_columns,
            "filter_applied": len(filters) > 0,
            "original_rows": len(data_rows),
            "filtered_rows": len(filtered_data_rows)
        }
        
        return {"results": results, "summary": summary}
    
    def _apply_filters_to_data(self, data_rows: List[List[Any]], headers: List[str], 
                              filters: List[Dict[str, Any]]) -> List[List[Any]]:
        """Apply filter conditions to data rows."""
        if not filters:
            return data_rows
        
        header_idx = {header: idx for idx, header in enumerate(headers)}
        filtered_rows = []
        
        for row in data_rows:
            include_row = True
            
            for filter_condition in filters:
                col = filter_condition.get("column")
                operator = filter_condition.get("operator", "=")
                value = filter_condition.get("value")
                logic = filter_condition.get("logic", "AND")
                
                if col not in header_idx:
                    continue
                
                col_idx = header_idx[col]
                if col_idx >= len(row):
                    continue
                
                row_value = row[col_idx]
                condition_met = self._evaluate_filter_condition(row_value, operator, value)
                
                if logic == "AND" and not condition_met:
                    include_row = False
                    break
                elif logic == "OR" and condition_met:
                    include_row = True
                    # For OR logic, if any condition is met, include the row
                    break
            
            if include_row:
                filtered_rows.append(row)
        
        return filtered_rows
    
    def _evaluate_filter_condition(self, row_value: Any, operator: str, filter_value: Any) -> bool:
        """Evaluate a single filter condition."""
        try:
            if operator == "=":
                return str(row_value).lower() == str(filter_value).lower()
            elif operator == "!=":
                return str(row_value).lower() != str(filter_value).lower()
            elif operator == ">":
                return float(row_value) > float(filter_value)
            elif operator == "<":
                return float(row_value) < float(filter_value)
            elif operator == ">=":
                return float(row_value) >= float(filter_value)
            elif operator == "<=":
                return float(row_value) <= float(filter_value)
            elif operator == "LIKE":
                return str(filter_value).lower() in str(row_value).lower()
            elif operator == "BETWEEN":
                # Expect filter_value to be "value1 AND value2"
                if " AND " in str(filter_value):
                    val1, val2 = str(filter_value).split(" AND ")
                    return float(val1) <= float(row_value) <= float(val2)
                return False
            elif operator == "IN":
                # Expect filter_value to be a list or comma-separated string
                if isinstance(filter_value, list):
                    return row_value in filter_value
                else:
                    values = [v.strip() for v in str(filter_value).split(",")]
                    return str(row_value) in values
        except (ValueError, TypeError):
            return False
        
        return False
    
    def _calculate_single_aggregation(self, values: List[float], aggregation: str) -> float:
        """Calculate a single aggregation on a list of values."""
        import statistics
        
        if not values:
            return 0
        
        if aggregation in ["SUM", "TOTAL"]:
            return sum(values)
        elif aggregation in ["AVERAGE", "MEAN", "AVG"]:
            return sum(values) / len(values)
        elif aggregation in ["MAX", "MAXIMUM"]:
            return max(values)
        elif aggregation in ["MIN", "MINIMUM"]:
            return min(values)
        elif aggregation in ["COUNT", "FREQUENCY"]:
            return len(values)
        elif aggregation == "MEDIAN":
            return statistics.median(values)
        elif aggregation in ["STDEV", "STD", "STANDARD_DEVIATION"]:
            return statistics.stdev(values) if len(values) > 1 else 0
        elif aggregation in ["VAR", "VARIANCE"]:
            return statistics.variance(values) if len(values) > 1 else 0
        elif aggregation == "MODE":
            try:
                return statistics.mode(values)
            except statistics.StatisticsError:
                return values[0]
        else:
            return sum(values)  # Default to SUM
    
    def _generate_multi_criteria_formula(self, aggregation: str, metric_column: str, 
                                       groupby_columns: List[str], group_values: Dict[str, Any],
                                       source_sheet: str, header_map: Dict[str, str], 
                                       result_row: int) -> str:
        """Generate Excel formula with multiple criteria (SUMIFS, AVERAGEIFS, COUNTIFS)."""
        
        # Map column names to Excel letters
        metric_col_letter = header_map.get(metric_column, metric_column) if header_map else metric_column
        
        # Build criteria ranges and criteria cells
        criteria_ranges = []
        criteria_cells = []
        
        for i, group_col in enumerate(groupby_columns):
            group_col_letter = header_map.get(group_col, group_col) if header_map else group_col
            criteria_ranges.append(f"{source_sheet}!{group_col_letter}:{group_col_letter}")
            
            # Reference cells containing the group values in the same row as the formula
            # Generate criteria cell references for current row
            criteria_cell = f"{chr(65 + i)}{result_row}"  # A4, B4, C4, etc. (same row as formula)

            criteria_cells.append(criteria_cell)
        
        # Generate formula based on aggregation type
        if aggregation in ["SUM", "TOTAL"]:
            # SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
            criteria_pairs = []
            for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                criteria_pairs.extend([range_ref, cell_ref])
            
            formula = f"=SUMIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{','.join(criteria_pairs)})"
            
        elif aggregation in ["AVERAGE", "MEAN", "AVG"]:
            # AVERAGEIFS(average_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
            criteria_pairs = []
            for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                criteria_pairs.extend([range_ref, cell_ref])
            
            formula = f"=AVERAGEIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{','.join(criteria_pairs)})"
            
        elif aggregation in ["COUNT", "FREQUENCY"]:
            # COUNTIFS(criteria_range1, criteria1, criteria_range2, criteria2, ...)
            criteria_pairs = []
            for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                criteria_pairs.extend([range_ref, cell_ref])
            
            formula = f"=COUNTIFS({','.join(criteria_pairs)})"
            
        elif aggregation in ["MAX", "MAXIMUM"]:
            # For multiple criteria, use array formula with IF conditions
            if len(groupby_columns) == 1:
                formula = f"=MAXIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{criteria_ranges[0]},{criteria_cells[0]})"
            else:
                # Build nested IF conditions
                conditions = []
                for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                    conditions.append(f"({range_ref}={cell_ref})")
                
                combined_condition = "*".join(conditions)
                formula = f"=MAX(IF({combined_condition},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
                
        elif aggregation in ["MIN", "MINIMUM"]:
            # Similar to MAX but with MIN
            if len(groupby_columns) == 1:
                formula = f"=MINIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{criteria_ranges[0]},{criteria_cells[0]})"
            else:
                conditions = []
                for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                    conditions.append(f"({range_ref}={cell_ref})")
                
                combined_condition = "*".join(conditions)
                formula = f"=MIN(IF({combined_condition},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
                
        else:
            # For other aggregations, use array formulas with IF conditions
            conditions = []
            for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                conditions.append(f"({range_ref}={cell_ref})")
            
            combined_condition = "*".join(conditions)
            
            if aggregation == "MEDIAN":
                formula = f"=MEDIAN(IF({combined_condition},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["STDEV", "STD", "STANDARD_DEVIATION"]:
                formula = f"=STDEV.P(IF({combined_condition},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            elif aggregation in ["VAR", "VARIANCE"]:
                formula = f"=VAR.P(IF({combined_condition},{source_sheet}!{metric_col_letter}:{metric_col_letter}))"
            else:
                # Fallback to SUMIFS for unknown aggregations
                criteria_pairs = []
                for range_ref, cell_ref in zip(criteria_ranges, criteria_cells):
                    criteria_pairs.extend([range_ref, cell_ref])
                formula = f"=SUMIFS({source_sheet}!{metric_col_letter}:{metric_col_letter},{','.join(criteria_pairs)})"
        
        return formula
    
    def _correct_formula_row_references(self, formula: str, target_row: int) -> str:
        """
        Correct formula row references to match the target row where the formula is placed.
        
        Recent fix: Ensures criteria references in formulas match the actual row where the formula appears.
        Fixes issues where formulas reference wrong rows (e.g., A63 instead of A58).
        
        Args:
            formula: The original formula with potentially incorrect row references
            target_row: The actual row number where the formula should be placed
            
        Returns:
            Corrected formula with proper row references
        """
        import re
        
        # Pattern to match criteria cell references (like A4, B4, C4) but not range references or sheet references
        # This matches: single letter + digits, but not ranges (A:A) or sheet references (Sheet1!)
        criteria_ref_pattern = r'(?<!:)(?<!!)([A-Z])(\d+)(?!:)'
        
        def replace_criteria_ref(match):
            column_letter = match.group(1)
            old_row = match.group(2)
            
            # Replace with the target row
            return f"{column_letter}{target_row}"
        
        # Replace only criteria cell references, preserving ranges like Sheet1!A:A
        corrected_formula = re.sub(criteria_ref_pattern, replace_criteria_ref, formula)
        
        return corrected_formula
    
    # Formula row correction: ensures criteria references match formula row placement
    
    def create_multi_aggregation_worksheet(self, multi_results: Dict[str, Any], 
                                         sheet_name: str = "Multi Aggregation",
                                         source_sheet: str = "Sheet1") -> Dict[str, Any]:
        """
        Create a worksheet from multi-aggregation results with proper formatting.
        
        Args:
            multi_results: Results from _calculate_multi_aggregation_grouped
            sheet_name: Name for the worksheet
            source_sheet: Source sheet name for formulas
            
        Returns:
            Worksheet dictionary with cells, formulas, and formatting
        """
        cells = []
        formulas = []
        row = 1
        
        results = multi_results.get("results", [])
        summary = multi_results.get("summary", {})
        
        if not results:
            # Empty results case
            cells.append({
                "address": "A1",
                "value": "No Results Found",
                "style": {"bold": True, "fontSize": 12, "color": "#D32F2F"}
            })
            return {"name": sheet_name, "cells": cells, "formulas": formulas}
        
        # Title and summary
        cells.append({
            "address": f"A{row}",
            "value": f"Multi-Aggregation Analysis",
            "style": {"bold": True, "fontSize": 14, "color": "#1976D2"}
        })
        row += 1
        
        cells.append({
            "address": f"A{row}",
            "value": f"Groups: {summary.get('total_groups', 0)} | Aggregations: {summary.get('aggregations_count', 0)} | Group By: {', '.join(summary.get('groupby_columns', []))}",
            "style": {"fontSize": 10, "color": "#666666"}
        })
        row += 2
        
        # Determine column structure
        if not results:
            return {"name": sheet_name, "cells": cells, "formulas": formulas}
        
        first_result = results[0]
        groupby_columns = list(first_result["group_values"].keys())
        aggregation_keys = list(first_result["aggregation_results"].keys())
        
        # Headers
        col_idx = 0
        
        # Group columns headers
        for group_col in groupby_columns:
            cells.append({
                "address": f"{chr(65 + col_idx)}{row}",
                "value": group_col.replace("_", " ").title(),
                "style": {"bold": True, "backgroundColor": "#E3F2FD", "borderBottom": "thick"}
            })
            col_idx += 1
        
        # Aggregation columns headers
        for agg_key in aggregation_keys:
            agg_info = first_result["aggregation_results"][agg_key]
            header_text = f"{agg_info['function']} of {agg_info['column']}"
            cells.append({
                "address": f"{chr(65 + col_idx)}{row}",
                "value": header_text,
                "style": {"bold": True, "backgroundColor": "#E8F5E8", "borderBottom": "thick"}
            })
            col_idx += 1
        
        # Row count header
        cells.append({
            "address": f"{chr(65 + col_idx)}{row}",
            "value": "Row Count",
            "style": {"bold": True, "backgroundColor": "#FFF3E0", "borderBottom": "thick"}
        })
        
        row += 1
        
        # Data rows
        for result_idx, result in enumerate(results):
            col_idx = 0
            bg_color = "#FAFAFA" if result_idx % 2 == 0 else "white"
            
            # Group values
            for group_col in groupby_columns:
                value = result["group_values"].get(group_col, "")
                cells.append({
                    "address": f"{chr(65 + col_idx)}{row}",
                    "value": value,
                    "style": {"backgroundColor": bg_color}
                })
                col_idx += 1
            
            # Aggregation results
            for agg_key in aggregation_keys:
                agg_result = result["aggregation_results"].get(agg_key, {})
                value = agg_result.get("value", 0)
                formula = agg_result.get("formula", "")
                
                # Format value based on aggregation type
                if isinstance(value, float):
                    if agg_result.get("function") in ["AVERAGE", "MEAN", "AVG"]:
                        formatted_value = f"{value:.2f}"
                    elif agg_result.get("function") in ["COUNT", "FREQUENCY"]:
                        formatted_value = int(value)
                    else:
                        formatted_value = f"{value:,.2f}"
                else:
                    formatted_value = value
                
                cells.append({
                    "address": f"{chr(65 + col_idx)}{row}",
                    "value": formatted_value,
                    "style": {"backgroundColor": bg_color, "textAlign": "right"}
                })
                
                # Add formula if available
                if formula:
                    # Recent fix: Update formula to reference the correct row for criteria
                    # Replace any hardcoded row references with the actual row number
                    corrected_formula = self._correct_formula_row_references(formula, row)
                    formulas.append({
                        "address": f"{chr(65 + col_idx)}{row}",
                        "formula": corrected_formula
                    })
                
                col_idx += 1
            
            # Row count
            cells.append({
                "address": f"{chr(65 + col_idx)}{row}",
                "value": result.get("row_count", 0),
                "style": {"backgroundColor": bg_color, "textAlign": "right", "fontSize": 9, "color": "#666666"}
            })
            
            row += 1
        
        # Summary section
        row += 2
        cells.append({
            "address": f"A{row}",
            "value": "Summary Statistics:",
            "style": {"bold": True, "fontSize": 11}
        })
        row += 1
        
        # Calculate summary statistics for numeric aggregations
        for agg_key in aggregation_keys:
            agg_info = first_result["aggregation_results"][agg_key]
            if agg_info["function"] not in ["COUNT", "FREQUENCY"]:
                values = [r["aggregation_results"][agg_key]["value"] for r in results if agg_key in r["aggregation_results"]]
                if values:
                    total = sum(values)
                    avg = total / len(values)
                    max_val = max(values)
                    min_val = min(values)
                    
                    cells.append({
                        "address": f"A{row}",
                        "value": f"{agg_info['function']} of {agg_info['column']}:",
                        "style": {"bold": True, "fontSize": 10}
                    })
                    
                    cells.append({
                        "address": f"B{row}",
                        "value": f"Total: {total:,.2f} | Avg: {avg:.2f} | Max: {max_val:,.2f} | Min: {min_val:,.2f}",
                        "style": {"fontSize": 10}
                    })
                    row += 1
        
        # Filter info if applicable
        if summary.get("filter_applied"):
            row += 1
            cells.append({
                "address": f"A{row}",
                "value": f"Note: Filters applied. Showing {summary.get('filtered_rows', 0)} of {summary.get('original_rows', 0)} total rows.",
                "style": {"fontSize": 9, "color": "#666666", "italic": True}
            })
        
        
        return {
            "name": sheet_name,
            "cells": cells,
            "formulas": formulas,
            "notes": f"Multi-aggregation analysis with {len(aggregation_keys)} calculations across {len(groupby_columns)} grouping dimensions"
        }
    
    def parse_multi_aggregation_query(self, query: str, available_columns: List[str]) -> Dict[str, Any]:
        """
        Parse natural language queries for multi-aggregation patterns.
        
        Example queries:
        - "sum qty and average revenue by sku_code and id"
        - "count orders and total sales by region and category where status = 'completed'"
        - "max price and min cost by product and supplier"
        
        Args:
            query: Natural language query
            available_columns: List of available column names for validation
            
        Returns:
            Multi-aggregation pattern dictionary or None if not a multi-aggregation query
        """
        import re
        
        query_lower = query.lower().strip()
        
        # Check if this looks like a multi-aggregation query
        multi_agg_indicators = [
            r"\b(sum|count|average|avg|max|min|total)\b.*\band\b.*\b(sum|count|average|avg|max|min|total)\b",
            r"\b(sum|count|average|avg|max|min|total)\b.*\,.*\b(sum|count|average|avg|max|min|total)\b"
        ]
        
        is_multi_agg = any(re.search(pattern, query_lower) for pattern in multi_agg_indicators)
        
        if not is_multi_agg:
            return None
        
        
        # Extract aggregations
        aggregations = []
        
        # Pattern: "function column" pairs
        agg_patterns = [
            r"\b(sum|total)\s+(\w+)",
            r"\b(count|frequency)\s+(?:of\s+)?(\w+)",
            r"\b(average|avg|mean)\s+(\w+)",
            r"\b(max|maximum)\s+(\w+)",
            r"\b(min|minimum)\s+(\w+)",
            r"\b(median)\s+(\w+)",
            r"\b(stdev|std)\s+(\w+)"
        ]
        
        for pattern in agg_patterns:
            matches = re.findall(pattern, query_lower)
            for func, col in matches:
                # Normalize function name
                func_normalized = {
                    "sum": "SUM", "total": "SUM",
                    "count": "COUNT", "frequency": "COUNT",
                    "average": "AVERAGE", "avg": "AVERAGE", "mean": "AVERAGE",
                    "max": "MAX", "maximum": "MAX",
                    "min": "MIN", "minimum": "MIN",
                    "median": "MEDIAN",
                    "stdev": "STDEV", "std": "STDEV"
                }.get(func, func.upper())
                
                # Recent fix: Use improved column matching algorithm
                matching_col, match_method = self._find_best_column_match(col, available_columns)
                logger.debug(f"Column matching: '{col}' -> '{matching_col}' via {match_method}")
                
                if matching_col:
                    aggregations.append({"column": matching_col, "function": func_normalized})
        
        # Extract groupby columns
        groupby_columns = []
        groupby_patterns = [
            r"by\s+([\w\s,and]+?)(?:\s+where|\s+order|\s+limit|$)",
            r"group\s+by\s+([\w\s,and]+?)(?:\s+where|\s+order|\s+limit|$)",
            r"grouped\s+by\s+([\w\s,and]+?)(?:\s+where|\s+order|\s+limit|$)"
        ]
        
        for pattern in groupby_patterns:
            match = re.search(pattern, query_lower)
            if match:
                groupby_text = match.group(1)
                # Split by common delimiters
                potential_columns = re.split(r'[,\s]+and[,\s]+|[,\s]+', groupby_text)
                
                for col_candidate in potential_columns:
                    col_candidate = col_candidate.strip()
                    if not col_candidate or col_candidate in ['and', 'by']:
                        continue
                    
                    # Recent fix: Use improved column matching algorithm
                    matching_col, match_method = self._find_best_column_match(col_candidate, available_columns)
                    logger.debug(f"GroupBy column matching: '{col_candidate}' -> '{matching_col}' via {match_method}")
                    
                    if matching_col and matching_col not in groupby_columns:
                        groupby_columns.append(matching_col)
                break
        
        # Extract filters (optional)
        filters = []
        filter_patterns = [
            r"where\s+(\w+)\s*(=|!=|>|<|>=|<=)\s*['\"]?([^'\"\s]+)['\"]?",
            r"(\w+)\s*(=|!=|>|<|>=|<=)\s*['\"]?([^'\"\s]+)['\"]?"
        ]
        
        for pattern in filter_patterns:
            matches = re.findall(pattern, query_lower)
            for col, operator, value in matches:
                # Skip if this looks like it's part of aggregation or groupby
                if any(col in agg["column"].lower() for agg in aggregations):
                    continue
                if any(col in gc.lower() for gc in groupby_columns):
                    continue
                
                # Recent fix: Use improved column matching algorithm
                matching_col, match_method = self._find_best_column_match(col, available_columns)
                logger.debug(f"Filter column matching: '{col}' -> '{matching_col}' via {match_method}")
                
                if matching_col:
                    filters.append({
                        "column": matching_col,
                        "operator": operator,
                        "value": value.strip("'\""),
                        "logic": "AND"
                    })
        
        # Validate we have the minimum required components
        if len(aggregations) < 2 or len(groupby_columns) < 1:
            return None
        
        pattern = {
            "aggregations": aggregations,
            "groupby_columns": groupby_columns,
            "filters": filters
        }
        
        return pattern
    
    def _find_best_column_match(self, query_column: str, available_columns: List[str]) -> tuple[str, str]:
        """
        Enhanced column matching algorithm with prioritized strategies.
        
        Recent fix: Implements robust column matching with:
        - Exact case-insensitive matching (highest priority)
        - Normalized matching with underscore/space/hyphen handling
        - Partial matching with minimum length requirements
        - Comprehensive logging for debugging
        
        Args:
            query_column: Column name extracted from query
            available_columns: List of available column names
            
        Returns:
            Tuple of (matched_column_name, match_method) or (None, "no_match")
        """
        if not query_column or not available_columns:
            return None, "no_match"
        
        query_col_clean = query_column.strip()
        
        # Strategy 1: Exact match (case insensitive)
        for available_col in available_columns:
            if query_col_clean.lower() == available_col.lower():
                return available_col, "exact_match"
        
        # Strategy 2: Normalized exact match (handle underscores, spaces, hyphens)
        query_normalized = self._normalize_column_name(query_col_clean)
        for available_col in available_columns:
            available_normalized = self._normalize_column_name(available_col)
            if query_normalized == available_normalized:
                return available_col, "normalized_exact_match"
        
        # Strategy 3: Substring match with validation (moved up for better priority)
        # Only if query column is ≥3 chars to avoid false positives
        if len(query_col_clean) >= 3:
            exact_substring_matches = []
            
            # Check if query is substring of available column
            for available_col in available_columns:
                if query_col_clean.lower() in available_col.lower():
                    # Additional validation: avoid matching common words
                    if not self._is_common_word(query_col_clean):
                        exact_substring_matches.append(available_col)
            
            # Check if available column is substring of query (for longer queries)
            if len(query_col_clean) >= 4:
                for available_col in available_columns:
                    if available_col.lower() in query_col_clean.lower() and len(available_col) >= 3:
                        if not self._is_common_word(available_col):
                            exact_substring_matches.append(available_col)
            
            # Return shortest match (most specific) if found
            if exact_substring_matches:
                best_match = min(exact_substring_matches, key=len)
                return best_match, "validated_substring_match"
        
        # Strategy 4: Partial match with similarity scoring (as fallback)
        # Only match if query term is substantial (≥3 chars) and matches significant portion
        if len(query_col_clean) >= 3:
            best_match = None
            best_score = 0
            
            for available_col in available_columns:
                score = self._calculate_column_match_score(query_col_clean, available_col)
                if score > best_score and score >= 0.6:  # Minimum 60% similarity
                    best_score = score
                    best_match = available_col
            
            if best_match:
                return best_match, f"partial_match_score_{best_score:.2f}"
        
        return None, "no_match"
    
    def _normalize_column_name(self, column_name: str) -> str:
        """
        Normalize column name by removing/replacing separators and converting to lowercase.
        
        Args:
            column_name: Original column name
            
        Returns:
            Normalized column name
        """
        import re
        
        # Convert to lowercase and replace separators with empty string
        normalized = column_name.lower()
        normalized = re.sub(r'[_\-\s]+', '', normalized)
        
        return normalized
    
    def _calculate_column_match_score(self, query_col: str, available_col: str) -> float:
        """
        Calculate similarity score between query column and available column.
        
        Args:
            query_col: Column name from query
            available_col: Available column name
            
        Returns:
            Similarity score between 0.0 and 1.0
        """
        import difflib
        
        # Normalize both for comparison
        query_norm = self._normalize_column_name(query_col)
        available_norm = self._normalize_column_name(available_col)
        
        # Use sequence matcher for similarity
        matcher = difflib.SequenceMatcher(None, query_norm, available_norm)
        base_score = matcher.ratio()
        
        # Boost score for exact substring matches after normalization
        if query_norm in available_norm or available_norm in query_norm:
            base_score = min(1.0, base_score + 0.2)
        
        # Boost score if query is significant portion of available column
        if len(query_norm) >= 3 and query_norm in available_norm:
            coverage = len(query_norm) / len(available_norm)
            if coverage >= 0.5:  # Query covers at least half the column name
                base_score = min(1.0, base_score + 0.1)
        
        return base_score
    
    def _is_common_word(self, word: str) -> bool:
        """
        Check if a word is too common to be a reliable column identifier.
        
        Args:
            word: Word to check
            
        Returns:
            True if word is considered too common
        """
        common_words = {
            'id', 'name', 'date', 'time', 'code', 'type', 'status', 'data', 'info',
            'number', 'value', 'count', 'total', 'sum', 'avg', 'max', 'min',
            'and', 'or', 'by', 'of', 'the', 'is', 'in', 'on', 'at', 'to'
        }
        
        return word.lower() in common_words

    # =========================
    # Multi-Table Analysis Suite
    # =========================

    def create_analysis_suite(self, data: Dict[str, Any], question: str, context: str = "") -> Dict[str, Any]:
        """
        Create multiple related tables from one query with coordinated analysis.
        
        Recent fix: Added multi-table analysis suite creation with:
        - Enhanced LLM-driven table planning
        - Automatic relationship detection
        - Cross-references between related tables
        - Integrated summary dashboard
        
        Args:
            data: Data context with pattern_sample and metadata
            question: User query requiring multi-table analysis
            context: Additional context for analysis
            
        Returns:
            Dict with coordinated sheets, relationships, and summary dashboard
        """
        logger.info(f"Creating analysis suite for question: {question}")
        
        try:
            # Step 1: Enhanced LLM analysis with multi-table awareness
            llm_structure = self._enhanced_llm_analysis_for_suite(data, question, context)
            
            # Step 2: Plan table relationships
            relationships = self.plan_table_relationships(llm_structure.get("sheets", []), question)
            
            # Step 3: Create coordinated sheets with cross-references
            sheets = []
            header_map = self.map_headers_to_letters(data)
            
            # Create individual analysis tables
            for i, sheet_plan in enumerate(llm_structure.get("sheets", []), 1):
                logger.info(f"Creating analysis table {i}: {sheet_plan.get('name', 'unnamed')}")
                sheet = self._create_coordinated_sheet(sheet_plan, data, header_map, relationships)
                sheets.append(sheet)
            
            # Step 4: Generate cross-references between related tables
            self._apply_cross_references(sheets, relationships)
            
            # Step 5: Create summary dashboard
            dashboard = self.create_summary_dashboard(sheets, llm_structure, question)
            sheets.insert(0, dashboard)
            
            result = {
                "sheets": sheets,
                "analysis_plan": llm_structure.get("analysis_plan", {}),
                "table_relationships": relationships,
                "cross_references": self._extract_cross_reference_summary(sheets, relationships),
                "metadata": {
                    "created": datetime.now().isoformat(),
                    "analysis_suite": True,
                    "total_tables": len(sheets) - 1,  # Exclude dashboard
                    "total_relationships": len(relationships),
                    "analysis_focus": llm_structure.get("analysis_plan", {}).get("analysis_focus", ""),
                    "question": question,
                    "context": context
                }
            }
            
            logger.info(f"Analysis suite created successfully with {len(sheets)} sheets and {len(relationships)} relationships")
            return result
            
        except Exception as e:
            logger.error(f"Failed to create analysis suite: {str(e)}")
            # Fallback to standard LLM-driven workbook
            return self.create_llm_driven_workbook(data, question, context)

    def plan_table_relationships(self, sheets: List[Dict[str, Any]], question: str) -> List[Dict[str, Any]]:
        """
        Determine how tables should connect based on shared dimensions and analytical flow.
        
        Args:
            sheets: List of sheet plans from LLM analysis
            question: Original user question for context
            
        Returns:
            List of relationship dictionaries with source, target, type, and linking info
        """
        logger.debug(f"Planning relationships for {len(sheets)} tables")
        relationships = []
        
        # Sort sheets by priority for relationship planning
        sorted_sheets = sorted(sheets, key=lambda s: {"high": 3, "medium": 2, "low": 1}.get(s.get("priority", "medium"), 2), reverse=True)
        
        for i, source_sheet in enumerate(sorted_sheets):
            for j, target_sheet in enumerate(sorted_sheets):
                if i != j:
                    relationship = self._identify_relationship(source_sheet, target_sheet, question)
                    if relationship:
                        relationships.append(relationship)
        
        # Remove duplicate relationships and optimize
        relationships = self._optimize_relationships(relationships)
        
        logger.info(f"Planned {len(relationships)} table relationships")
        return relationships

    def generate_cross_references(self, source_sheet: Dict[str, Any], target_sheet: Dict[str, Any], 
                                 relationship: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Create cell references between related tables.
        
        Args:
            source_sheet: Source sheet data
            target_sheet: Target sheet data  
            relationship: Relationship definition between sheets
            
        Returns:
            List of cross-reference formulas to add to target sheet
        """
        logger.debug(f"Generating cross-references from {source_sheet.get('name')} to {target_sheet.get('name')}")
        cross_refs = []
        
        relationship_type = relationship.get("type")
        source_name = source_sheet.get("name", "Sheet1")
        linking_columns = relationship.get("linking_columns", [])
        
        if relationship_type == "summary_to_detail":
            # Create references from summary to detailed analysis
            cross_refs.extend(self._create_summary_detail_references(source_name, target_sheet, linking_columns))
            
        elif relationship_type == "cross_analysis":
            # Create comparison references between peer tables
            cross_refs.extend(self._create_cross_analysis_references(source_name, target_sheet, linking_columns))
            
        elif relationship_type == "drill_down":
            # Create drill-down references for hierarchical analysis
            cross_refs.extend(self._create_drill_down_references(source_name, target_sheet, linking_columns))
        
        logger.debug(f"Generated {len(cross_refs)} cross-references")
        return cross_refs

    def create_summary_dashboard(self, sheets: List[Dict[str, Any]], llm_structure: Dict[str, Any], 
                               question: str) -> Dict[str, Any]:
        """
        Combine multiple tables into an executive overview dashboard.
        
        Args:
            sheets: List of analysis sheets to summarize
            llm_structure: Original LLM analysis structure
            question: Original user question
            
        Returns:
            Dashboard sheet with summary of all analysis tables
        """
        logger.info("Creating summary dashboard")
        
        cells = []
        formulas = []
        row = 1
        
        # Dashboard title and overview
        cells.extend([
            {"address": f"A{row}", "value": "📊 Analysis Suite Dashboard"},
            {"address": f"A{row+1}", "value": f"Query: {question}"},
            {"address": f"A{row+2}", "value": f"Analysis Focus: {llm_structure.get('analysis_plan', {}).get('analysis_focus', 'Multi-table analysis')}"}
        ])
        row += 4
        
        # Key insights section
        key_insights = llm_structure.get("analysis_plan", {}).get("key_insights_to_find", [])
        if key_insights:
            cells.append({"address": f"A{row}", "value": "🔍 Key Insights to Find:"})
            row += 1
            for insight in key_insights:
                cells.append({"address": f"B{row}", "value": f"• {insight}"})
                row += 1
            row += 1
        
        # Tables overview section
        cells.append({"address": f"A{row}", "value": "📋 Analysis Tables:"})
        row += 1
        
        # Headers for table overview
        cells.extend([
            {"address": f"A{row}", "value": "Table Name"},
            {"address": f"B{row}", "value": "Purpose"},
            {"address": f"C{row}", "value": "Priority"},
            {"address": f"D{row}", "value": "Key Metrics"},
            {"address": f"E{row}", "value": "Quick Links"}
        ])
        row += 1
        
        # Summary of each analysis table
        for sheet in sheets:
            if sheet.get("name") != "Dashboard":  # Skip dashboard itself
                sheet_name = sheet.get("name", "Unknown")
                
                # Get metrics summary
                formulas_list = sheet.get("formulas", [])
                metrics_count = len(formulas_list)
                metrics_summary = f"{metrics_count} metrics" if metrics_count > 0 else "No metrics"
                
                cells.extend([
                    {"address": f"A{row}", "value": sheet_name},
                    {"address": f"B{row}", "value": sheet.get("notes", "Analysis table")},
                    {"address": f"C{row}", "value": "High"},  # Default priority
                    {"address": f"D{row}", "value": metrics_summary}
                ])
                
                # Create hyperlink to sheet (Excel formula)
                formulas.append({
                    "address": f"E{row}",
                    "formula": f'=HYPERLINK("#\'{sheet_name}\'!A1","→ Go to {sheet_name}")',
                    "description": f"Link to {sheet_name} sheet"
                })
                row += 1
        
        row += 2
        
        # Executive summary with key metrics from other sheets
        cells.append({"address": f"A{row}", "value": "📈 Executive Summary:"})
        row += 1
        
        # Add cross-sheet summary formulas
        summary_formulas = self._create_executive_summary_formulas(sheets, row)
        formulas.extend(summary_formulas)
        
        # Add cells for summary labels
        for i, formula in enumerate(summary_formulas):
            cells.append({
                "address": f"A{row + i}",
                "value": formula.get("description", f"Summary Metric {i+1}")
            })
        
        return _build_worksheet_payload(
            name="Dashboard",
            cells=cells,
            formulas=formulas,
            notes="Executive dashboard summarizing all analysis tables with cross-references and key insights"
        )

    # =========================
    # Enhanced LLM Integration
    # =========================

    def _enhanced_llm_analysis_for_suite(self, data: Dict[str, Any], question: str, context: str) -> Dict[str, Any]:
        """
        Enhanced LLM analysis specifically for multi-table suite creation.
        Extends existing llm_decide_metrics_and_structure() with relationship awareness.
        """
        logger.debug("Running enhanced LLM analysis for multi-table suite")
        
        # First get standard LLM analysis
        base_structure = self.llm_decide_metrics_and_structure(data, question, context)
        
        # Enhance with multi-table awareness if we have multiple sheets
        sheets = base_structure.get("sheets", [])
        if len(sheets) > 1:
            # Add relationship metadata to each sheet
            for i, sheet in enumerate(sheets):
                sheet["suite_index"] = i
                sheet["relationship_candidates"] = self._identify_relationship_candidates(sheet, sheets)
                
            # Enhance analysis plan with suite-specific information
            analysis_plan = base_structure.get("analysis_plan", {})
            analysis_plan["multi_table_suite"] = True
            analysis_plan["table_count"] = len(sheets)
            analysis_plan["coordination_strategy"] = self._determine_coordination_strategy(sheets, question)
            
            base_structure["analysis_plan"] = analysis_plan
        
        logger.debug(f"Enhanced LLM analysis completed with {len(sheets)} tables")
        return base_structure

    # =========================
    # Relationship Planning Helpers
    # =========================

    def _identify_relationship(self, source_sheet: Dict[str, Any], target_sheet: Dict[str, Any], 
                             question: str) -> Optional[Dict[str, Any]]:
        """Identify the relationship type between two sheets."""
        source_metrics = source_sheet.get("metrics", [])
        target_metrics = target_sheet.get("metrics", [])
        
        # Extract columns used in metrics
        source_columns = set()
        target_columns = set()
        
        for metric in source_metrics:
            if "target_column" in metric:
                source_columns.add(metric["target_column"])
        
        for metric in target_metrics:
            if "target_column" in metric:
                target_columns.add(metric["target_column"])
        
        # Find common columns (potential linking columns)
        common_columns = list(source_columns & target_columns)
        
        if not common_columns:
            return None
        
        # Determine relationship type based on sheet characteristics
        source_priority = source_sheet.get("priority", "medium")
        target_priority = target_sheet.get("priority", "medium")
        
        if source_priority == "high" and target_priority in ["medium", "low"]:
            relationship_type = "summary_to_detail"
        elif len(source_metrics) < len(target_metrics):
            relationship_type = "drill_down"
        else:
            relationship_type = "cross_analysis"
        
        return {
            "source": source_sheet.get("name"),
            "target": target_sheet.get("name"),
            "type": relationship_type,
            "linking_columns": common_columns,
            "strength": len(common_columns) / max(len(source_columns), len(target_columns), 1),
            "description": f"{relationship_type.replace('_', ' ').title()} relationship via {', '.join(common_columns)}"
        }

    def _optimize_relationships(self, relationships: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Remove duplicate and weak relationships."""
        # Remove duplicates based on source-target pairs
        seen = set()
        unique_relationships = []
        
        for rel in relationships:
            key = (rel["source"], rel["target"])
            reverse_key = (rel["target"], rel["source"])
            
            if key not in seen and reverse_key not in seen:
                seen.add(key)
                unique_relationships.append(rel)
        
        # Keep only strong relationships (strength > 0.3)
        strong_relationships = [rel for rel in unique_relationships if rel.get("strength", 0) > 0.3]
        
        return strong_relationships

    def _identify_relationship_candidates(self, sheet: Dict[str, Any], all_sheets: List[Dict[str, Any]]) -> List[str]:
        """Identify potential relationship candidates for a sheet."""
        candidates = []
        sheet_columns = set()
        
        # Extract columns from this sheet's metrics
        for metric in sheet.get("metrics", []):
            if "target_column" in metric:
                sheet_columns.add(metric["target_column"])
        
        # Find other sheets with overlapping columns
        for other_sheet in all_sheets:
            if other_sheet.get("name") != sheet.get("name"):
                other_columns = set()
                for metric in other_sheet.get("metrics", []):
                    if "target_column" in metric:
                        other_columns.add(metric["target_column"])
                
                if sheet_columns & other_columns:  # Has common columns
                    candidates.append(other_sheet.get("name"))
        
        return candidates

    def _determine_coordination_strategy(self, sheets: List[Dict[str, Any]], question: str) -> str:
        """Determine the overall coordination strategy for the suite."""
        if len(sheets) <= 2:
            return "simple_comparison"
        elif any("summary" in sheet.get("purpose", "").lower() for sheet in sheets):
            return "hierarchical_summary"
        elif "trend" in question.lower() or "time" in question.lower():
            return "temporal_analysis"
        else:
            return "cross_dimensional_analysis"

    # =========================
    # Cross-Reference Generation
    # =========================

    def _create_coordinated_sheet(self, sheet_plan: Dict[str, Any], data: Dict[str, Any], 
                                header_map: Dict[str, str], relationships: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create a sheet with awareness of its relationships to other sheets."""
        # Start with standard sheet creation
        sheet = self._create_sheet_from_llm_plan(sheet_plan, data, header_map)
        
        # Add relationship-aware enhancements
        sheet_name = sheet_plan.get("name")
        related_relationships = [rel for rel in relationships if rel.get("source") == sheet_name or rel.get("target") == sheet_name]
        
        if related_relationships:
            # Add cross-reference section
            cross_ref_cells = []
            cross_ref_formulas = []
            
            # Find current row count to add cross-references below existing content
            existing_cells = sheet.get("cells", [])
            max_row = 0
            for cell in existing_cells:
                addr = cell.get("address", "A1")
                # Extract row number from address like "A5" -> 5
                row_match = re.search(r'(\d+)', addr)
                if row_match:
                    max_row = max(max_row, int(row_match.group(1)))
            
            start_row = max_row + 3  # Leave some space
            
            # Add cross-reference header
            cross_ref_cells.append({
                "address": f"A{start_row}",
                "value": "🔗 Related Tables:"
            })
            
            # Add references to related tables
            for i, rel in enumerate(related_relationships):
                row = start_row + i + 1
                if rel.get("source") == sheet_name:
                    related_table = rel.get("target")
                    relation_desc = f"→ {related_table} ({rel.get('type', 'related')})"
                else:
                    related_table = rel.get("source")
                    relation_desc = f"← {related_table} ({rel.get('type', 'related')})"
                
                cross_ref_cells.append({
                    "address": f"A{row}",
                    "value": relation_desc
                })
                
                # Add hyperlink to related sheet
                cross_ref_formulas.append({
                    "address": f"B{row}",
                    "formula": f'=HYPERLINK("#\'{related_table}\'!A1","Go to {related_table}")',
                    "description": f"Link to {related_table}"
                })
            
            # Merge cross-references into sheet
            sheet["cells"].extend(cross_ref_cells)
            sheet["formulas"].extend(cross_ref_formulas)
        
        return sheet

    def _apply_cross_references(self, sheets: List[Dict[str, Any]], relationships: List[Dict[str, Any]]):
        """Apply cross-references between sheets based on relationships."""
        for relationship in relationships:
            source_name = relationship.get("source")
            target_name = relationship.get("target")
            
            # Find source and target sheets
            source_sheet = next((s for s in sheets if s.get("name") == source_name), None)
            target_sheet = next((s for s in sheets if s.get("name") == target_name), None)
            
            if source_sheet and target_sheet:
                # Generate cross-references
                cross_refs = self.generate_cross_references(source_sheet, target_sheet, relationship)
                
                # Add cross-references to target sheet
                target_sheet["formulas"].extend(cross_refs)

    def _create_summary_detail_references(self, source_name: str, target_sheet: Dict[str, Any], 
                                        linking_columns: List[str]) -> List[Dict[str, Any]]:
        """Create references from summary to detail tables."""
        references = []
        
        # Create a reference to the summary value
        references.append({
            "address": "F1",
            "formula": f"='{source_name}'!B3",
            "description": f"Summary reference from {source_name}"
        })
        
        return references

    def _create_cross_analysis_references(self, source_name: str, target_sheet: Dict[str, Any], 
                                        linking_columns: List[str]) -> List[Dict[str, Any]]:
        """Create comparison references between peer tables."""
        references = []
        
        # Create comparison formula
        references.append({
            "address": "G1",
            "formula": f"='{source_name}'!B3-B3",
            "description": f"Comparison with {source_name}"
        })
        
        return references

    def _create_drill_down_references(self, source_name: str, target_sheet: Dict[str, Any], 
                                    linking_columns: List[str]) -> List[Dict[str, Any]]:
        """Create drill-down references for hierarchical analysis."""
        references = []
        
        # Create drill-down validation
        references.append({
            "address": "H1",
            "formula": f"=SUM(B:B)-'{source_name}'!B3",
            "description": f"Drill-down validation against {source_name}"
        })
        
        return references

    def _extract_cross_reference_summary(self, sheets: List[Dict[str, Any]], 
                                       relationships: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Extract summary of all cross-references for metadata."""
        summary = {
            "total_references": 0,
            "reference_types": [],
            "connected_sheets": set()
        }
        
        for sheet in sheets:
            formulas = sheet.get("formulas", [])
            for formula in formulas:
                if "'" in formula.get("formula", ""):  # Contains sheet reference
                    summary["total_references"] += 1
                    summary["connected_sheets"].add(sheet.get("name"))
        
        summary["connected_sheets"] = list(summary["connected_sheets"])
        summary["reference_types"] = list(set(rel.get("type") for rel in relationships))
        
        return summary

    def _create_executive_summary_formulas(self, sheets: List[Dict[str, Any]], start_row: int) -> List[Dict[str, Any]]:
        """Create executive summary formulas that aggregate data from all sheets."""
        summary_formulas = []
        
        # Count total analysis tables
        summary_formulas.append({
            "address": f"B{start_row}",
            "formula": f"={len([s for s in sheets if s.get('name') != 'Dashboard'])}",
            "description": "Total Analysis Tables"
        })
        
        # Create aggregation formulas if we have numeric sheets
        numeric_sheets = [s for s in sheets if s.get("formulas") and s.get("name") != "Dashboard"]
        if numeric_sheets:
            # Sum of first metrics across sheets
            sheet_refs = []
            for sheet in numeric_sheets[:3]:  # Limit to first 3 sheets to avoid formula complexity
                sheet_name = sheet.get("name")
                sheet_refs.append(f"'{sheet_name}'!B3")
            
            if sheet_refs:
                summary_formulas.append({
                    "address": f"B{start_row + 1}",
                    "formula": f"={'+'.join(sheet_refs)}",
                    "description": "Cross-Table Sum"
                })
        
        #return summary_formulas


# =========================
# Integration Example and Testing
# =========================

def create_multi_table_workbook(data: Dict[str, Any], question: str, context: str = "") -> Dict[str, Any]:
    """
    Convenience function that demonstrates the new multi-table analysis suite functionality.
    
    This function integrates all the new methods:
    1. create_analysis_suite() - creates multiple related tables
    2. plan_table_relationships() - determines table connections  
    3. generate_cross_references() - creates cell references
    4. create_summary_dashboard() - combines into overview
    
    Args:
        data: Data context with pattern_sample and metadata
        question: User query requiring multi-table analysis
        context: Additional context for analysis
        
    Returns:
        Complete workbook with coordinated multi-table analysis
        
    Example:
        >>> data = {
        ...     'pattern_sample': [
        ...         ['region', 'product', 'sales', 'quantity'],
        ...         ['North', 'A', 1000, 50],
        ...         ['South', 'B', 1500, 75]
        ...     ]
        ... }
        >>> query = "sum sales and average quantity by region and product"
        >>> result = create_multi_table_workbook(data, query)
        >>> print(f"Created {len(result['sheets'])} coordinated sheets")
    """
    try:
        manager = WorksheetManager()
        
        # Use the new analysis suite functionality
        result = manager.create_analysis_suite(data, question, context)
        
        # Enhance with additional metadata
        result["integration_demo"] = {
            "methods_used": [
                "create_analysis_suite",
                "plan_table_relationships", 
                "generate_cross_references",
                "create_summary_dashboard"
            ],
            "enhanced_llm_integration": True,
            "coordinated_analysis": True
        }
        
        return result
        
    except Exception as e:
        logger.error(f"Multi-table workbook creation failed: {str(e)}")
        # Fallback to standard analysis
        manager = WorksheetManager()
        return manager.create_llm_driven_workbook(data, question, context)


# =========================
# Multi-Table Analysis Documentation
# =========================

"""
MULTI-TABLE ANALYSIS SUITE - IMPLEMENTATION SUMMARY

The WorksheetManager class has been extended with comprehensive multi-table analysis capabilities:

CORE METHODS ADDED:
1. create_analysis_suite() - Main entry point for multi-table analysis
2. plan_table_relationships() - Determines table connections and dependencies  
3. generate_cross_references() - Creates Excel formulas linking related tables
4. create_summary_dashboard() - Executive overview combining all analyses

ENHANCED LLM INTEGRATION:
- _enhanced_llm_analysis_for_suite() extends existing llm_decide_metrics_and_structure()
- Adds relationship awareness and coordination strategy to LLM analysis
- Maintains full backward compatibility with existing analysis methods

RELATIONSHIP TYPES SUPPORTED:
- summary_to_detail: High-level summaries feeding detailed breakdowns
- cross_analysis: Peer-level tables with shared dimensions
- drill_down: Hierarchical analysis with increasing granularity

CROSS-REFERENCE CAPABILITIES:
- Automatic Excel hyperlinks between related sheets
- Summary validation formulas (drill-down totals match summary)
- Comparison formulas between peer analyses
- Executive dashboard aggregations across all tables

USAGE EXAMPLES:
```python
# Basic multi-table analysis
manager = WorksheetManager()
result = manager.create_analysis_suite(data, "sum sales and average profit by region and product")

# Convenience function with full integration
result = create_multi_table_workbook(data, "complex multi-dimensional query")

# Access relationships and cross-references
relationships = result["table_relationships"]
cross_refs = result["cross_references"]
```

INTEGRATION WITH EXISTING PATTERNS:
- Fully compatible with existing _build_worksheet_payload() structure
- Uses existing header mapping and formula generation patterns
- Extends but doesn't modify core LLM analysis functionality
- Maintains all existing logging and error handling patterns

USAGE EXAMPLES:

# Enhanced LLM Formula Generation
manager = WorksheetManager()
result = manager.generate_dynamic_formula(
    query_context="calculate weighted average price by region with error handling",
    data_context=data_with_samples,
    force_llm=True
)


# Backward Compatible Usage
traditional_result = manager._generate_formula_from_metric(
    metric={"formula_type": "SUM", "target_column": "sales"},
    header_map={"sales": "C"},
    source_sheet="Sheet1"
)
"""