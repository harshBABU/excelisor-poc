# Enhanced Intent Classification System

## Overview

The `intent.py` module has been completely revamped to provide robust, LLM-enhanced intent classification for Excel assistant operations. This system intelligently understands user queries, extracts relevant entities, and makes informed recommendations for data analysis.

## Key Features

### 🚀 **LLM-Powered Classification**
- Uses OpenAI GPT-4 for sophisticated intent understanding
- Dynamic prompting with data context awareness
- Graceful fallback to rule-based classification when LLM is unavailable

### 🧠 **Intelligent Entity Extraction**
- Automatically identifies metrics, dimensions, aggregations, and operations
- Extracts time references, comparisons, and numerical thresholds
- Context-aware entity recognition based on data columns

### 📊 **Smart Column Selection**
- Semantic understanding of column relevance to user queries
- Validates suggestions against actual DataFrame columns
- Provides reasoning for column selection decisions

### 🎯 **Enhanced Action Types**
- Extended from 4 to 11 action types for better classification granularity
- Includes: analyze, audit, chart, formula, filter, pivot, sort, summary, compare, forecast
- Adaptive confidence scoring with detailed confidence levels

### 🔄 **Backward Compatibility**
- Maintains existing API for seamless integration
- Original functions work exactly as before
- New enhanced functions available for advanced usage

## Architecture

### Core Classes

#### `ActionType` (Enum)
Enhanced action types for Excel operations:
- `ANALYZE`: General data exploration and insights
- `AUDIT`: Data quality checks, missing values, duplicates
- `CHART`: Create visualizations (charts, graphs, plots)
- `FORMULA`: Create calculations or Excel formulas
- `FILTER`: Filter/subset data based on conditions
- `PIVOT`: Create pivot tables or grouped summaries
- `SORT`: Sort or rank data
- `SUMMARY`: Statistical summaries and descriptions
- `COMPARE`: Compare different groups or time periods
- `FORECAST`: Predictive analysis or trend projection
- `UNKNOWN`: Cannot determine intent

#### `DataComplexity` (Enum)
Data complexity assessment:
- `SIMPLE`: < 100 cells
- `MEDIUM`: 100-5000 cells
- `LARGE`: 5000-50000 cells
- `MASSIVE`: > 50000 cells

#### `ConfidenceLevel` (Enum)
Confidence assessment levels:
- `VERY_LOW`: 0.0-0.3
- `LOW`: 0.3-0.5
- `MEDIUM`: 0.5-0.7
- `HIGH`: 0.7-0.9
- `VERY_HIGH`: 0.9-1.0

#### `EntityExtraction` (Dataclass)
Extracted entities from user queries:
```python
@dataclass
class EntityExtraction:
    metrics: List[str]           # sales, profit, quantity
    dimensions: List[str]        # product, region, date
    aggregations: List[str]      # sum, avg, max
    time_references: List[str]   # last month, Q1, yearly
    comparisons: List[str]       # vs, compared to, better than
    thresholds: List[float]      # > 1000, < 50%
    operations: List[str]        # filter, sort, rank
```

#### `IntentResult` (Dataclass)
Comprehensive classification result:
```python
@dataclass
class IntentResult:
    action_type: ActionType
    confidence: float
    confidence_level: ConfidenceLevel
    data_complexity: DataComplexity
    entities: EntityExtraction
    reasoning: str
    suggested_columns: Dict[str, List[str]]
    analysis_plan: Optional[Dict[str, Any]] = None
    llm_enhanced: bool = False
    fallback_used: bool = False
```

### Main Classifier

#### `EnhancedIntentClassifier`
The core classification engine that:
1. **Initializes OpenAI client** (if available)
2. **Assesses data complexity** based on DataFrame size
3. **Attempts LLM classification** with sophisticated prompting
4. **Falls back to rule-based** classification if LLM fails
5. **Validates and enhances** all results

## API Reference

### Backward Compatible Functions

#### `classify_intent(question: str, df: pd.DataFrame) -> Dict[str, Any]`
Original function signature maintained for backward compatibility.

**Returns:**
```python
{
    "action_type": str,      # One of the action types
    "confidence": float,     # 0.0-1.0
    "data_complexity": str   # simple/medium/large/massive
}
```

#### `select_relevant_columns(question: str, header_analysis: Dict[str, Any]) -> Dict[str, Any]`
Enhanced column selection with LLM reasoning.

**Returns:**
```python
{
    "numeric_cols": List[str],
    "categorical_cols": List[str],
    "reasoning": str,
    "aggregation_pattern": Optional[str],
    "groupby_pattern": Optional[Dict]
}
```

### Enhanced Functions

#### `classify_intent_enhanced(question: str, df: pd.DataFrame, context: Optional[str] = None) -> IntentResult`
Full-featured intent classification with comprehensive results.

#### `extract_query_entities(question: str, df: pd.DataFrame) -> EntityExtraction`
Extract entities from user query using LLM or pattern matching.

#### `get_analysis_recommendations(question: str, df: pd.DataFrame) -> Dict[str, Any]`
Get comprehensive analysis recommendations based on intent classification.

## Usage Examples

### Basic Usage (Backward Compatible)
```python
import pandas as pd
from intent import classify_intent, select_relevant_columns

# Basic intent classification
df = pd.DataFrame({'Sales': [100, 200, 300], 'Region': ['A', 'B', 'C']})
result = classify_intent("analyze sales by region", df)
print(result)
# {'action_type': 'analyze', 'confidence': 0.9, 'data_complexity': 'simple'}

# Column selection
header_analysis = {
    'Sales': {'is_numeric': True, 'is_categorical': False},
    'Region': {'is_numeric': False, 'is_categorical': True}
}
columns = select_relevant_columns("sum sales by region", header_analysis)
print(columns['numeric_cols'])  # ['Sales']
print(columns['categorical_cols'])  # ['Region']
```

### Enhanced Usage
```python
from intent import classify_intent_enhanced, extract_query_entities, get_analysis_recommendations

# Enhanced classification with full details
result = classify_intent_enhanced("show me the top 10 products by sales", df)
print(f"Action: {result.action_type.value}")
print(f"Confidence: {result.confidence}")
print(f"Reasoning: {result.reasoning}")
print(f"Suggested columns: {result.suggested_columns}")

# Entity extraction
entities = extract_query_entities("sum sales by region for products > 1000", df)
print(f"Metrics: {entities.metrics}")
print(f"Aggregations: {entities.aggregations}")
print(f"Thresholds: {entities.thresholds}")

# Comprehensive recommendations
recommendations = get_analysis_recommendations("create a chart showing trends", df)
print(f"Primary action: {recommendations['primary_action']}")
print(f"Recommended charts: {recommendations['analysis_plan']['recommended_charts']}")
```

## Configuration

### Environment Variables
- `OPENAI_API_KEY`: Required for LLM-enhanced features
- If not set, system automatically falls back to rule-based classification

### Dependencies
- `openai>=1.3.7`: For LLM functionality
- `pandas>=2.3.1`: For data handling
- `pydantic>=2.12.0`: For data validation

## Error Handling

The system includes comprehensive error handling:

1. **OpenAI API Failures**: Automatic fallback to rule-based classification
2. **Invalid DataFrames**: Safe handling with informative error messages  
3. **Malformed Queries**: Graceful degradation with fallback logic
4. **Network Issues**: Timeout and retry mechanisms built into OpenAI client

## Performance Considerations

### LLM Mode
- **Latency**: ~1-3 seconds per classification (OpenAI API call)
- **Accuracy**: Very high for complex queries
- **Cost**: ~$0.001-0.003 per classification

### Fallback Mode  
- **Latency**: ~1-10ms per classification
- **Accuracy**: Good for common patterns
- **Cost**: Free

### Recommendations
- Use LLM mode for production with complex user queries
- Use fallback mode for development or simple patterns
- Consider caching results for repeated queries

## Migration Guide

### From Old System
The new system is fully backward compatible. Existing code will work without changes:

```python
# This continues to work exactly as before
result = classify_intent("analyze sales", df)
columns = select_relevant_columns("sum by region", header_analysis)
```

### To Enhanced System
To leverage new capabilities, gradually migrate to enhanced functions:

```python
# Replace this:
result = classify_intent("analyze sales", df)

# With this:
result = classify_intent_enhanced("analyze sales", df)
action = result.action_type.value  # Get the action type
confidence = result.confidence     # Get confidence
reasoning = result.reasoning       # Get detailed reasoning
entities = result.entities         # Get extracted entities
```

## Testing

The system includes comprehensive tests covering:
- ✅ LLM classification success and failure scenarios
- ✅ Fallback classification for all action types
- ✅ Entity extraction and validation
- ✅ Column selection and validation
- ✅ Error handling and edge cases
- ✅ Backward compatibility

Run tests with: `python test_intent.py`

## Future Enhancements

### Planned Features
1. **Conversation Memory**: Multi-turn context awareness
2. **Custom Prompts**: User-configurable classification prompts
3. **Batch Processing**: Efficient classification of multiple queries
4. **Model Selection**: Support for different LLM models
5. **Caching Layer**: Results caching for performance optimization

### Integration Opportunities
1. **Feedback Loop**: Learn from user corrections
2. **Domain Adaptation**: Industry-specific classification
3. **Multilingual Support**: Non-English query understanding
4. **Advanced Analytics**: Query complexity analysis and recommendations

## Support

For questions, issues, or contributions:
1. Check the comprehensive test suite for usage examples
2. Review error logs for troubleshooting guidance  
3. Consult this documentation for API reference
4. Consider the backward compatibility functions for simple use cases

The enhanced intent system represents a significant improvement in robustness, accuracy, and capabilities while maintaining full backward compatibility with existing code.
