# Formula Requirements Extraction System

## Overview

The `extract_formula_requirements()` method in the `EnhancedIntentClassifier` provides intelligent parsing of natural language queries to extract structured formula requirements for dynamic Excel formula generation.

## Key Features

### 🎯 **Comprehensive Extraction**
- **Target Columns**: Identifies specific data columns mentioned or implied in queries
- **Aggregation Functions**: Detects calculations (SUM, AVERAGE, COUNT, etc.)
- **Grouping Requirements**: Extracts GROUP BY patterns and specifications
- **Filter Conditions**: Parses WHERE clauses and conditional logic
- **Sort Criteria**: Identifies ORDER BY requirements and ranking needs
- **Additional Requirements**: Detects complex logic, custom formulas, and special functions

### 🧠 **Hybrid Intelligence**
- **LLM-Powered**: Uses OpenAI GPT-4o-mini for advanced natural language understanding
- **Rule-Based Fallback**: Comprehensive regex-based extraction when LLM is unavailable
- **Smart Validation**: Validates column names against actual DataFrame columns
- **Fuzzy Matching**: Finds similar column names when exact matches aren't found

### 📊 **Query Complexity Assessment**
- **Simple**: Basic single-column operations
- **Medium**: Multi-column operations with basic grouping
- **Complex**: Advanced operations with multiple conditions and custom logic

## API Reference

### Main Method

```python
def extract_formula_requirements(self, question: str, df: pd.DataFrame, 
                               context: Optional[str] = None) -> Dict[str, Any]
```

#### Parameters
- `question`: User's natural language query
- `df`: DataFrame containing the data for column validation
- `context`: Optional conversation context for better understanding

#### Returns
```python
{
    "target_columns": List[str],     # Columns to calculate on
    "aggregation_functions": List[str],  # Functions to apply
    "grouping_requirements": {       # Groupby specifications
        "group_by_columns": List[str],
        "group_by_type": str,        # "simple|complex|nested"
        "having_conditions": List[str]
    },
    "filter_conditions": [           # WHERE clauses
        {
            "column": str,
            "operator": str,         # "=|>|<|>=|<=|!=|LIKE|BETWEEN|IN"
            "value": str,
            "logic": str            # "AND|OR"
        }
    ],
    "sort_criteria": {              # ORDER BY specifications
        "sort_columns": List[str],
        "sort_directions": List[str], # ["ASC|DESC"]
        "limit": Optional[int]
    },
    "additional_requirements": {     # Complex logic detection
        "conditional_logic": bool,
        "custom_formulas": List[str],
        "date_functions": bool,
        "text_functions": bool,
        "lookup_functions": bool,
        "complex_calculations": bool
    },
    "confidence": float,            # Extraction confidence (0.0-1.0)
    "query_complexity": str,        # "simple|medium|complex"
    "extraction_method": str,       # "llm|rule_based|default_fallback"
    "validation": {                 # Validation metadata
        "columns_validated": bool,
        "available_columns": List[str],
        "validation_timestamp": str
    }
}
```

## Usage Examples

### Basic Usage

```python
from intent import get_intent_classifier
import pandas as pd

# Initialize classifier
classifier = get_intent_classifier()

# Sample data
df = pd.DataFrame({
    'sku_code': ['A001', 'A002', 'B001'],
    'sales_amount': [1000, 1500, 800],
    'region': ['North', 'South', 'North'],
    'date': ['2024-01-01', '2024-01-02', '2024-01-03']
})

# Extract requirements
requirements = classifier.extract_formula_requirements(
    "Sum sales amount by region for top 5 SKUs",
    df
)

print(requirements)
```

### Example Extractions

#### 1. Simple Aggregation
**Query**: "Sum sales amount"
```python
{
    "target_columns": ["sales_amount"],
    "aggregation_functions": ["SUM"],
    "grouping_requirements": {"group_by_columns": [], "group_by_type": "simple"},
    "filter_conditions": [],
    "sort_criteria": {"sort_columns": [], "sort_directions": [], "limit": None},
    "query_complexity": "simple",
    "confidence": 0.9
}
```

#### 2. Grouped Aggregation
**Query**: "Average sales by region"
```python
{
    "target_columns": ["sales_amount"],
    "aggregation_functions": ["AVERAGE"],
    "grouping_requirements": {
        "group_by_columns": ["region"],
        "group_by_type": "simple"
    },
    "filter_conditions": [],
    "sort_criteria": {"sort_columns": [], "sort_directions": [], "limit": None},
    "query_complexity": "medium",
    "confidence": 0.85
}
```

#### 3. Complex Query with Filters and Sorting
**Query**: "Top 5 SKUs by sales where region = 'North' ordered by sales descending"
```python
{
    "target_columns": ["sales_amount"],
    "aggregation_functions": ["SUM"],
    "grouping_requirements": {
        "group_by_columns": ["sku_code"],
        "group_by_type": "simple"
    },
    "filter_conditions": [
        {
            "column": "region",
            "operator": "=",
            "value": "North",
            "logic": "AND"
        }
    ],
    "sort_criteria": {
        "sort_columns": ["sales_amount"],
        "sort_directions": ["DESC"],
        "limit": 5
    },
    "query_complexity": "complex",
    "confidence": 0.8
}
```

#### 4. Conditional Logic
**Query**: "Count SKUs if sales amount greater than 1000"
```python
{
    "target_columns": ["sku_code"],
    "aggregation_functions": ["COUNTIF"],
    "grouping_requirements": {"group_by_columns": [], "group_by_type": "simple"},
    "filter_conditions": [
        {
            "column": "sales_amount",
            "operator": ">",
            "value": "1000",
            "logic": "AND"
        }
    ],
    "additional_requirements": {
        "conditional_logic": true,
        "complex_calculations": false
    },
    "query_complexity": "medium",
    "confidence": 0.82
}
```

## Supported Aggregation Functions

### Basic Aggregations
- **SUM/TOTAL**: Sum values
- **AVERAGE/MEAN**: Calculate average
- **COUNT**: Count occurrences
- **MAX/MAXIMUM**: Find maximum value
- **MIN/MINIMUM**: Find minimum value
- **MEDIAN**: Calculate median

### Statistical Functions
- **STDEV**: Standard deviation
- **VAR**: Variance
- **MODE**: Most frequent value
- **PERCENTILE**: Calculate percentiles
- **QUARTILE**: Calculate quartiles

### Conditional Functions
- **SUMIF/SUMIFS**: Conditional sum
- **COUNTIF/COUNTIFS**: Conditional count
- **AVERAGEIF/AVERAGEIFS**: Conditional average

### Text Functions
- **CONCATENATE**: Join text
- **LEFT/RIGHT/MID**: Text extraction
- **UPPER/LOWER**: Case conversion

### Date Functions
- **YEAR/MONTH/DAY**: Date extraction
- **WEEKDAY**: Day of week
- **NOW/TODAY**: Current date/time

### Lookup Functions
- **VLOOKUP/XLOOKUP**: Vertical lookup
- **INDEX/MATCH**: Index matching

## Grouping Pattern Detection

The system recognizes various grouping patterns:

- `"by [column]"` → GROUP BY column
- `"for each [category]"` → GROUP BY category
- `"per [dimension]"` → GROUP BY dimension
- `"broken down by"` → GROUP BY
- `"grouped by"` → GROUP BY

## Filter Pattern Detection

Supports multiple filter patterns:

- `"where [condition]"` → WHERE clause
- `"for [category] = [value]"` → WHERE category = value
- `"if [condition]"` → conditional logic
- `"greater than"`, `"less than"`, `"equal to"` → comparison operators
- `"between [x] and [y]"` → range conditions
- `"contains"`, `"starts with"`, `"ends with"` → text matching

## Sort Pattern Detection

Recognizes sorting requirements:

- `"order by"`, `"sort by"` → ORDER BY
- `"top [N]"`, `"bottom [N]"` → LIMIT with sort
- `"highest"`, `"lowest"` → descending/ascending sort
- `"ascending"`, `"descending"` → sort direction

## Integration with Dynamic Formula Generation

The extracted requirements can be seamlessly integrated with the enhanced formula generation system:

```python
# Extract requirements
requirements = classifier.extract_formula_requirements(query, df)

# Generate formula using requirements
from worksheet_manager import WorksheetManager
manager = WorksheetManager()

enhanced_formula = manager.generate_enhanced_formula(
    query=query,
    data_context={"columns": df.dtypes.to_dict()},
    formula_requirements=requirements
)
```

## Error Handling and Fallbacks

The system provides robust error handling:

1. **LLM Failure**: Automatically falls back to rule-based extraction
2. **Validation Errors**: Attempts fuzzy column matching
3. **Complete Failure**: Returns sensible defaults with low confidence
4. **Column Mismatch**: Suggests similar column names

## Performance Considerations

- **LLM Calls**: Cached and optimized for performance
- **Rule-Based**: Fast fallback for simple queries
- **Validation**: Efficient column matching algorithms
- **Memory**: Lightweight data structures for requirements

## Best Practices

1. **Provide Context**: Include conversation context for better understanding
2. **Clear Queries**: Use specific column names when possible
3. **Validate Results**: Check confidence scores and validation metadata
4. **Handle Fallbacks**: Implement graceful degradation for low-confidence results
5. **Test Edge Cases**: Verify behavior with complex queries

