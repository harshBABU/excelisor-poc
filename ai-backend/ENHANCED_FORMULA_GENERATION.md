# Enhanced Formula Generation with LLM Integration

## Overview

The enhanced formula generation system combines traditional template-based Excel formulas with intelligent LLM-powered formula creation for complex queries that go beyond standard aggregations.

## Key Features

### 1. **Hybrid Approach**
- **Template-based**: Fast, reliable formulas for standard operations (SUM, AVERAGE, COUNT, etc.)
- **LLM-powered**: Intelligent formula generation for complex business logic, multiple criteria, and custom calculations
- **Automatic fallback**: Graceful degradation to templates when LLM is unavailable

### 2. **Smart Complexity Detection**
The system automatically determines when to use LLM generation based on:
- Multiple conditions or criteria
- Complex aggregations (weighted, conditional, nested)
- Time-based calculations
- Statistical operations not in templates
- Custom business logic
- Multi-step calculations

### 3. **Comprehensive Error Handling**
- Validation of LLM-generated formulas
- Automatic fallback to template methods
- Formula syntax checking and cleaning
- Proper sheet reference validation

## Usage Examples

### Basic Usage
```python
# Initialize WorksheetManager
wm = WorksheetManager()

# Example 1: Simple aggregation (uses templates)
result = wm.generate_enhanced_formula(
    query="sum of sales column",
    data_context=data_context
)
# Returns: =SUM(Sheet1!A:A)

# Example 2: Complex conditional logic (uses LLM)
result = wm.generate_enhanced_formula(
    query="calculate weighted average of revenue by region where sales > 1000 and date is within last 30 days",
    data_context=data_context
)
# Returns: Custom LLM-generated formula with complex logic

# Example 3: Multi-criteria aggregation (uses LLM)
result = wm.generate_enhanced_formula(
    query="sum revenue for products in category A or B where quantity > 10 and discount < 0.2",
    data_context=data_context
)
# Returns: =SUMIFS(...) with multiple criteria
```

### Advanced Usage with Custom Requirements
```python
# Specify custom requirements
requirements = {
    "formula_type": "SUMIFS",
    "target_column": "revenue",
    "criteria_columns": ["category", "quantity", "discount"],
    "criteria_values": ["A", ">10", "<0.2"],
    "use_llm_generation": True
}

result = wm.generate_enhanced_formula(
    query="sum revenue with multiple conditions",
    data_context=data_context,
    formula_requirements=requirements
)
```

### Integration with Existing Workflows
```python
# In make_calculated_values_sheet method
def make_calculated_values_sheet(self, data: Dict[str, Any], hint: str = ""):
    # ... existing code ...
    
    # Enhanced formula generation for complex queries
    if self._is_complex_query(hint):
        enhanced_result = self.generate_enhanced_formula(
            query=hint,
            data_context=data,
            source_sheet=source_sheet
        )
        
        if enhanced_result["success"]:
            # Use LLM-generated formula
            formula = enhanced_result["formula"]
            method_used = enhanced_result["method_used"]
            logger.info(f"Using {method_used} formula: {formula}")
        else:
            # Fallback to existing logic
            formula = self._generate_formula_from_metric(metric, header_map)
```

## Result Structure

The `generate_enhanced_formula` method returns a comprehensive result dictionary:

```python
{
    "formula": "=SUMIFS(Sheet1!C:C,Sheet1!A:A,\"A\",Sheet1!B:B,\">10\")",
    "method_used": "llm",  # "llm", "template", or "fallback"
    "requirements": {
        "formula_type": "SUMIFS",
        "target_column": "C",
        "has_conditions": True,
        "use_llm_generation": True
    },
    "source_sheet": "Sheet1",
    "success": True,
    "error": None
}
```

## Complexity Detection Criteria

The system uses LLM generation when queries contain:

### Multiple Conditions
- "multiple", "and where", "or where"
- Example: "sum sales where category is A and region is North"

### Complex Aggregations
- "weighted", "conditional", "nested"
- Example: "weighted average of prices by volume"

### Time-Based Calculations
- "date", "time", "period", "monthly", "yearly"
- Example: "sum sales for last 3 months"

### Statistical Operations
- "correlation", "regression", "trend"
- Example: "calculate correlation between price and sales"

### Custom Business Logic
- "if then", "rate", "ratio", "calculate"
- Example: "calculate commission rate based on sales tier"

### Multi-Step Calculations
- "first then", "step"
- Example: "first calculate discount then apply tax"

## Error Handling and Fallbacks

### 1. LLM Unavailable
```python
# Automatic fallback to template method
if not self.openai_client:
    logger.warning("OpenAI client not available, falling back to template-based generation")
    return self._fallback_to_template_formula(metric_requirements, source_sheet)
```

### 2. LLM Generation Fails
```python
# Exception handling with graceful degradation
try:
    llm_formula = self._generate_llm_formula(...)
    return llm_formula
except Exception as e:
    logger.warning(f"LLM formula generation failed, falling back to template: {str(e)}")
    # Falls back to template method
```

### 3. Formula Validation
```python
# Validates and cleans LLM output
validated_formula = self._validate_and_clean_formula(formula, data_context, source_sheet)
# Checks for:
# - Proper = prefix
# - Balanced parentheses
# - Sheet references
# - Removes markdown formatting
```

## Configuration and Setup

### Prerequisites
```python
# OpenAI client must be configured in WorksheetManager
wm = WorksheetManager()
# Requires self.openai_client to be properly initialized

# Data context must include:
data_context = {
    "header_analysis": {...},  # Column types and sample values
    "pattern_sample": [...],   # Sample data rows
    # ... other context data
}
```

### Environment Variables
```bash
# Required for LLM functionality
OPENAI_API_KEY=your_api_key_here
```

## Performance Considerations

### Template vs LLM Usage
- **Templates**: ~1ms response time, deterministic results
- **LLM**: ~1-3s response time, intelligent but variable results
- **Automatic selection**: Balances speed and capability

### Caching Strategy
Consider implementing caching for similar queries:
```python
# Future enhancement - formula caching
formula_cache = {}
cache_key = f"{query_hash}_{data_structure_hash}"
if cache_key in formula_cache:
    return formula_cache[cache_key]
```

## Best Practices

### 1. Query Formulation
```python
# Good: Specific and clear
"sum revenue for products where category is electronics and price > 100"

# Avoid: Vague or ambiguous
"calculate something with sales data"
```

### 2. Data Context
```python
# Ensure comprehensive data context
data_context = {
    "header_analysis": complete_column_info,
    "pattern_sample": representative_sample,
    "metadata": relevant_business_context
}
```

### 3. Error Handling
```python
# Always check result success
result = wm.generate_enhanced_formula(query, data_context)
if result["success"]:
    use_formula(result["formula"])
else:
    handle_error(result["error"])
```

## Future Enhancements

### 1. Formula Optimization
- Implement formula performance analysis
- Suggest more efficient alternatives
- Cache commonly used formulas

### 2. Business Rule Integration
- Store domain-specific formula patterns
- Learn from user corrections and preferences
- Integrate with business glossary

### 3. Multi-Language Support
- Support queries in multiple languages
- Localized Excel function names
- Regional formula formatting

## Troubleshooting

### Common Issues

1. **LLM Returns Invalid Formula**
   - Solution: Validation catches most issues, falls back to templates
   - Check: Formula syntax and sheet references

2. **Performance Concerns**
   - Monitor LLM usage patterns
   - Implement caching for repeated queries
   - Use complexity detection to limit LLM calls

3. **API Rate Limits**
   - Implement exponential backoff
   - Cache results for similar queries
   - Use template fallbacks when rate limited

### Debug Logging
```python
# Enable debug logging to trace formula generation
import logging
logging.getLogger('worksheet_manager').setLevel(logging.DEBUG)
```

