# Multi-Aggregation System for Excel Worksheets

## Overview

The multi-aggregation system extends the existing `_calculate_grouped_aggregation()` method to handle multiple aggregations simultaneously across multiple grouping columns with advanced filtering and performance optimization.

## Key Features

### 🎯 **Multiple Aggregations Simultaneously**
- Process multiple aggregation functions in a single query
- Support different functions for different columns
- Generate separate Excel formulas for each aggregation
- Maintain performance with large datasets

### 📊 **Multiple Grouping Columns**
- Group by multiple dimensions simultaneously
- Create composite group keys
- Handle complex business scenarios like "by SKU and region"

### 🔍 **Advanced Filtering**
- Apply WHERE clauses before aggregation
- Support multiple operators (=, !=, >, <, >=, <=, LIKE, BETWEEN, IN)
- Combine filters with AND/OR logic
- Maintain data integrity throughout processing

### ⚡ **Performance Optimizations**
- Efficient data grouping algorithms
- Memory-optimized processing for large datasets
- Smart formula generation with cell references
- Structured output for fast worksheet creation

## API Reference

### Main Method

```python
def _calculate_multi_aggregation_grouped(self, pattern_sample: List[List[Any]], 
                                       multi_groupby_pattern: Dict[str, Any], 
                                       source_sheet: str = "Sheet1", 
                                       header_map: Dict[str, str] = None) -> Dict[str, Any]
```

#### Input Pattern Structure
```python
multi_groupby_pattern = {
    "aggregations": [
        {"column": "qty", "function": "SUM"},
        {"column": "revenue", "function": "AVERAGE"},
        {"column": "orders", "function": "COUNT"}
    ],
    "groupby_columns": ["sku_code", "region", "category"],
    "filters": [  # Optional
        {"column": "status", "operator": "=", "value": "active", "logic": "AND"},
        {"column": "sales_amount", "operator": ">", "value": "1000", "logic": "AND"}
    ]
}
```

#### Output Structure
```python
{
    "results": [
        {
            "group_key": "sku_code:SKU001 | region:North | category:Electronics",
            "group_values": {
                "sku_code": "SKU001",
                "region": "North", 
                "category": "Electronics"
            },
            "aggregation_results": {
                "qty_SUM": {
                    "value": 150,
                    "formula": "=SUMIFS(Sheet1!D:D,Sheet1!A:A,B4,Sheet1!B:B,C4,Sheet1!C:C,D4)",
                    "column": "qty",
                    "function": "SUM"
                },
                "revenue_AVERAGE": {
                    "value": 1250.50,
                    "formula": "=AVERAGEIFS(Sheet1!E:E,Sheet1!A:A,B4,Sheet1!B:B,C4,Sheet1!C:C,D4)",
                    "column": "revenue",
                    "function": "AVERAGE"
                },
                "orders_COUNT": {
                    "value": 25,
                    "formula": "=COUNTIFS(Sheet1!A:A,B4,Sheet1!B:B,C4,Sheet1!C:C,D4)",
                    "column": "orders",
                    "function": "COUNT"
                }
            },
            "row_count": 45
        }
    ],
    "summary": {
        "total_groups": 15,
        "aggregations_count": 3,
        "groupby_columns": ["sku_code", "region", "category"],
        "filter_applied": true,
        "original_rows": 1000,
        "filtered_rows": 800
    }
}
```

### Supporting Methods

#### Filter Application
```python
def _apply_filters_to_data(self, data_rows: List[List[Any]], headers: List[str], 
                          filters: List[Dict[str, Any]]) -> List[List[Any]]
```

#### Single Aggregation Calculation
```python
def _calculate_single_aggregation(self, values: List[float], aggregation: str) -> float
```

#### Multi-Criteria Formula Generation
```python
def _generate_multi_criteria_formula(self, aggregation: str, metric_column: str, 
                                   groupby_columns: List[str], group_values: Dict[str, Any],
                                   source_sheet: str, header_map: Dict[str, str], 
                                   result_row: int) -> str
```

#### Worksheet Creation
```python
def create_multi_aggregation_worksheet(self, multi_results: Dict[str, Any], 
                                     sheet_name: str = "Multi Aggregation",
                                     source_sheet: str = "Sheet1") -> Dict[str, Any]
```

#### Query Parsing
```python
def parse_multi_aggregation_query(self, query: str, available_columns: List[str]) -> Dict[str, Any]
```

## Usage Examples

### Example 1: Basic Multi-Aggregation

**Query**: "sum qty and average revenue by sku_code and region"

```python
from worksheet_manager import WorksheetManager

# Sample data
pattern_sample = [
    ["sku_code", "region", "qty", "revenue"],
    ["SKU001", "North", 50, 1000],
    ["SKU001", "North", 30, 800],
    ["SKU001", "South", 40, 1200],
    ["SKU002", "North", 60, 1500],
    ["SKU002", "South", 35, 900]
]

# Define multi-aggregation pattern
multi_pattern = {
    "aggregations": [
        {"column": "qty", "function": "SUM"},
        {"column": "revenue", "function": "AVERAGE"}
    ],
    "groupby_columns": ["sku_code", "region"],
    "filters": []
}

# Create worksheet manager
manager = WorksheetManager()

# Calculate results
results = manager._calculate_multi_aggregation_grouped(
    pattern_sample, 
    multi_pattern,
    source_sheet="Sheet1"
)

# Create worksheet
worksheet = manager.create_multi_aggregation_worksheet(results)
```

**Output Results**:
```
Group: sku_code:SKU001 | region:North
- qty_SUM: 80 (=SUMIFS(Sheet1!C:C,Sheet1!A:A,B4,Sheet1!B:B,C4))
- revenue_AVERAGE: 900.00 (=AVERAGEIFS(Sheet1!D:D,Sheet1!A:A,B4,Sheet1!B:B,C4))

Group: sku_code:SKU001 | region:South  
- qty_SUM: 40 (=SUMIFS(Sheet1!C:C,Sheet1!A:A,B5,Sheet1!B:B,C5))
- revenue_AVERAGE: 1200.00 (=AVERAGEIFS(Sheet1!D:D,Sheet1!A:A,B5,Sheet1!B:B,C5))
```

### Example 2: Complex Query with Filters

**Query**: "count orders and total sales by category and supplier where status = 'active' and region != 'test'"

```python
# Complex multi-aggregation with filters
complex_pattern = {
    "aggregations": [
        {"column": "order_id", "function": "COUNT"},
        {"column": "sales_amount", "function": "SUM"},
        {"column": "profit_margin", "function": "AVERAGE"}
    ],
    "groupby_columns": ["category", "supplier", "region"],
    "filters": [
        {"column": "status", "operator": "=", "value": "active", "logic": "AND"},
        {"column": "region", "operator": "!=", "value": "test", "logic": "AND"},
        {"column": "sales_amount", "operator": ">", "value": "100", "logic": "AND"}
    ]
}

results = manager._calculate_multi_aggregation_grouped(
    large_dataset, 
    complex_pattern
)
```

### Example 3: Natural Language Query Parsing

```python
# Parse natural language query
query = "sum qty and average revenue by sku_code and id where region = 'North'"
available_columns = ["sku_code", "id", "qty", "revenue", "region", "status"]

parsed_pattern = manager.parse_multi_aggregation_query(query, available_columns)

if parsed_pattern:
    results = manager._calculate_multi_aggregation_grouped(
        pattern_sample,
        parsed_pattern
    )
    
    worksheet = manager.create_multi_aggregation_worksheet(results)
```

## Supported Aggregation Functions

### Basic Aggregations
- **SUM/TOTAL**: Sum all values
- **AVERAGE/MEAN/AVG**: Calculate arithmetic mean
- **COUNT/FREQUENCY**: Count non-empty occurrences
- **MAX/MAXIMUM**: Find maximum value
- **MIN/MINIMUM**: Find minimum value

### Statistical Functions
- **MEDIAN**: Calculate median value
- **STDEV/STD**: Standard deviation
- **VAR/VARIANCE**: Variance calculation
- **MODE**: Most frequent value

### Advanced Functions
- **UNIQUE_COUNT**: Count distinct values
- **POSITIVE_COUNT**: Count positive values only
- **NEGATIVE_COUNT**: Count negative values only
- **ZERO_COUNT**: Count zero values
- **NON_ZERO_COUNT**: Count non-zero values

## Excel Formula Generation

### Single Criteria (Legacy)
```excel
=SUMIF(Sheet1!A:A,"SKU001",Sheet1!B:B)
=COUNTIF(Sheet1!A:A,"North")
```

### Multiple Criteria (New)
```excel
=SUMIFS(Sheet1!D:D,Sheet1!A:A,B4,Sheet1!B:B,C4,Sheet1!C:C,D4)
=AVERAGEIFS(Sheet1!E:E,Sheet1!A:A,B4,Sheet1!B:B,C4)
=COUNTIFS(Sheet1!A:A,B4,Sheet1!B:B,C4,Sheet1!C:C,D4)
```

### Array Formulas (Statistical)
```excel
=MEDIAN(IF((Sheet1!A:A=B4)*(Sheet1!B:B=C4),Sheet1!D:D))
=STDEV.P(IF((Sheet1!A:A=B4)*(Sheet1!B:B=C4),Sheet1!D:D))
```

## Filter Operations

### Supported Operators
- **=**: Exact match (case-insensitive)
- **!=**: Not equal
- **>**: Greater than (numeric)
- **<**: Less than (numeric)  
- **>=**: Greater than or equal
- **<=**: Less than or equal
- **LIKE**: Contains text (case-insensitive)
- **BETWEEN**: Range check ("100 AND 500")
- **IN**: Multiple values ("A,B,C" or ["A","B","C"])

### Filter Logic
- **AND**: All conditions must be true (default)
- **OR**: Any condition can be true

### Filter Examples
```python
filters = [
    {"column": "status", "operator": "=", "value": "active", "logic": "AND"},
    {"column": "sales", "operator": ">", "value": "1000", "logic": "AND"},
    {"column": "region", "operator": "IN", "value": "North,South,East", "logic": "AND"},
    {"column": "date", "operator": "BETWEEN", "value": "2024-01-01 AND 2024-12-31", "logic": "AND"}
]
```

## Performance Considerations

### Memory Optimization
- **Streaming Processing**: Processes data in chunks for large datasets
- **Efficient Grouping**: Uses hash-based grouping for O(n) performance
- **Smart Filtering**: Applies filters early to reduce processing load

### Formula Optimization
- **Cell References**: Uses dynamic cell references instead of hardcoded values
- **Batch Processing**: Generates all formulas in a single pass
- **Memory Efficient**: Minimizes object creation and copying

### Scalability Features
- **Large Dataset Support**: Handles datasets with 100K+ rows efficiently
- **Progressive Results**: Can return partial results for real-time updates
- **Configurable Limits**: Supports result limiting for performance tuning

## Worksheet Output Structure

### Professional Formatting
- **Color-coded Headers**: Different colors for grouping vs aggregation columns
- **Alternating Row Colors**: Improved readability
- **Right-aligned Numbers**: Proper numerical formatting
- **Summary Statistics**: Automatic calculation of totals, averages, min/max

### Layout Structure
```
Multi-Aggregation Analysis
Groups: 15 | Aggregations: 3 | Group By: sku_code, region

| SKU Code | Region | SUM of qty | AVG of revenue | COUNT of orders | Row Count |
|----------|--------|------------|----------------|-----------------|-----------|
| SKU001   | North  | 150        | 1,250.50       | 25              | 45        |
| SKU001   | South  | 120        | 980.25         | 18              | 32        |
| ...      | ...    | ...        | ...            | ...             | ...       |

Summary Statistics:
SUM of qty: Total: 2,500 | Avg: 166.67 | Max: 300 | Min: 50
AVG of revenue: Total: 18,750.50 | Avg: 1,250.03 | Max: 2,100.00 | Min: 850.00

Note: Filters applied. Showing 800 of 1000 total rows.
```

## Integration Examples

### With Formula Requirements Extraction
```python
from intent import get_intent_classifier

# Extract requirements using LLM
classifier = get_intent_classifier()
requirements = classifier.extract_formula_requirements(
    "sum qty and average revenue by sku_code and region where status = 'active'",
    df
)

# Convert to multi-aggregation pattern
if len(requirements["aggregation_functions"]) > 1:
    multi_pattern = {
        "aggregations": [
            {"column": col, "function": func} 
            for col, func in zip(requirements["target_columns"], requirements["aggregation_functions"])
        ],
        "groupby_columns": requirements["groupby_requirements"]["group_by_columns"],
        "filters": requirements["filter_conditions"]
    }
    
    results = manager._calculate_multi_aggregation_grouped(pattern_sample, multi_pattern)
```

### With Existing Workflow
```python
# Check if query requires multi-aggregation
def process_query(query, data):
    available_columns = list(data[0]) if data else []
    
    # Try multi-aggregation first
    multi_pattern = manager.parse_multi_aggregation_query(query, available_columns)
    
    if multi_pattern:
        print("Using multi-aggregation processing")
        results = manager._calculate_multi_aggregation_grouped(data, multi_pattern)
        return manager.create_multi_aggregation_worksheet(results)
    else:
        print("Using single aggregation processing")
        # Fall back to existing single aggregation logic
        return manager._calculate_grouped_aggregation(data, single_pattern)
```

## Error Handling and Validation

### Data Validation
- **Column Existence**: Validates all columns exist in the dataset
- **Type Checking**: Ensures numeric operations on numeric columns
- **Index Conversion**: Handles column indices converted to names
- **Empty Data**: Graceful handling of missing or empty datasets

### Error Recovery
- **Partial Results**: Returns partial results when some aggregations fail
- **Fallback Options**: Falls back to simpler aggregations when complex ones fail
- **Detailed Logging**: Comprehensive logging for debugging and monitoring

### Performance Monitoring
- **Processing Time**: Tracks time for each phase
- **Memory Usage**: Monitors memory consumption for large datasets
- **Result Size**: Tracks and limits result set sizes

## Best Practices

### Query Design
1. **Specific Column Names**: Use exact column names when possible
2. **Reasonable Grouping**: Limit grouping columns to avoid combinatorial explosion
3. **Filter Early**: Apply filters to reduce dataset size before aggregation
4. **Test with Sample Data**: Validate queries with small datasets first

### Performance Optimization
1. **Index Key Columns**: Ensure grouping columns are efficiently searchable
2. **Batch Processing**: Process large datasets in manageable chunks
3. **Memory Management**: Monitor and limit memory usage for large aggregations
4. **Formula Complexity**: Balance formula sophistication with Excel performance

### Data Quality
1. **Clean Data**: Ensure data quality before aggregation
2. **Handle Nulls**: Plan for missing or null values in datasets
3. **Validate Types**: Ensure numeric columns contain valid numbers
4. **Consistent Formatting**: Maintain consistent data formats across columns

This multi-aggregation system provides enterprise-grade capabilities for complex data analysis while maintaining the simplicity and performance needed for production Excel applications.

