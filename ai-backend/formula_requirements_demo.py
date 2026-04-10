#!/usr/bin/env python3
"""
Demo script for the Formula Requirements Extraction system.
Shows how to use the extract_formula_requirements() method.
"""

import pandas as pd
import json
from intent import get_intent_classifier

def demo_formula_requirements_extraction():
    """Demonstrate the formula requirements extraction functionality."""
    
    print("🧮 Formula Requirements Extraction Demo")
    print("=" * 50)
    
    # Create sample data
    sample_data = {
        'id': ['A001', 'A002', 'B001', 'B002', 'C001'],
        'sku_code': ['SKU-001', 'SKU-002', 'SKU-003', 'SKU-004', 'SKU-005'],
        'sales_amount': [1000, 1500, 800, 2000, 1200],
        'quantity': [10, 15, 8, 20, 12],
        'region': ['North', 'South', 'North', 'West', 'South'],
        'category': ['Electronics', 'Clothing', 'Electronics', 'Home', 'Clothing'],
        'date': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04', '2024-01-05']
    }
    
    df = pd.DataFrame(sample_data)
    print(f"Sample Data ({len(df)} rows):")
    print(df.head())
    print()
    
    # Initialize classifier
    classifier = get_intent_classifier()
    
    # Test queries with increasing complexity
    test_queries = [
        {
            "query": "Sum sales amount",
            "description": "Simple aggregation"
        },
        {
            "query": "Average sales by region",
            "description": "Grouped aggregation"
        },
        {
            "query": "Count SKUs by category",
            "description": "Count with grouping"
        },
        {
            "query": "Top 3 SKUs by sales amount",
            "description": "Ranking query"
        },
        {
            "query": "Sum sales where region = 'North'",
            "description": "Filtered aggregation"
        },
        {
            "query": "Average quantity by category for sales greater than 1000",
            "description": "Complex query with filter and grouping"
        },
        {
            "query": "Total sales by region ordered by sales descending",
            "description": "Aggregation with sorting"
        },
        {
            "query": "Count of SKUs if sales amount between 1000 and 1500",
            "description": "Conditional count with range"
        }
    ]
    
    for i, test_case in enumerate(test_queries, 1):
        print(f"📝 Test {i}: {test_case['description']}")
        print(f"Query: \"{test_case['query']}\"")
        print("-" * 40)
        
        try:
            # Extract requirements
            requirements = classifier.extract_formula_requirements(
                test_case['query'], 
                df
            )
            
            # Display key results
            print(f"✅ Extraction successful!")
            print(f"   Confidence: {requirements['confidence']:.2f}")
            print(f"   Complexity: {requirements['query_complexity']}")
            print(f"   Method: {requirements.get('extraction_method', 'unknown')}")
            print()
            
            # Show extracted components
            if requirements['target_columns']:
                print(f"🎯 Target Columns: {', '.join(requirements['target_columns'])}")
            
            if requirements['aggregation_functions']:
                print(f"📊 Aggregations: {', '.join(requirements['aggregation_functions'])}")
            
            if requirements['grouping_requirements']['group_by_columns']:
                print(f"📑 Group By: {', '.join(requirements['grouping_requirements']['group_by_columns'])}")
            
            if requirements['filter_conditions']:
                filters = []
                for condition in requirements['filter_conditions']:
                    filters.append(f"{condition['column']} {condition['operator']} {condition['value']}")
                print(f"🔍 Filters: {'; '.join(filters)}")
            
            if requirements['sort_criteria']['sort_columns']:
                sorts = []
                for col, direction in zip(
                    requirements['sort_criteria']['sort_columns'],
                    requirements['sort_criteria']['sort_directions']
                ):
                    sorts.append(f"{col} {direction}")
                print(f"🔄 Sort: {'; '.join(sorts)}")
                
                if requirements['sort_criteria']['limit']:
                    print(f"📏 Limit: {requirements['sort_criteria']['limit']}")
            
            # Show additional requirements
            additional = requirements['additional_requirements']
            active_features = [k for k, v in additional.items() if v and k != 'custom_formulas']
            if active_features:
                print(f"⚡ Additional: {', '.join(active_features)}")
            
            print()
            
            # Optionally show full JSON for complex queries
            if requirements['query_complexity'] == 'complex':
                print("📋 Full Requirements (JSON):")
                print(json.dumps(requirements, indent=2, default=str))
                print()
                
        except Exception as e:
            print(f"❌ Extraction failed: {str(e)}")
            print()
        
        print("=" * 50)
        print()

def demo_integration_with_formula_generation():
    """Show how requirements integrate with formula generation."""
    
    print("🔗 Integration with Formula Generation")
    print("=" * 50)
    
    # Sample data
    df = pd.DataFrame({
        'sku': ['A001', 'A002', 'B001'],
        'sales': [1000, 1500, 800],
        'region': ['North', 'South', 'North']
    })
    
    classifier = get_intent_classifier()
    
    query = "Sum sales by region"
    requirements = classifier.extract_formula_requirements(query, df)
    
    print(f"Query: \"{query}\"")
    print()
    print("Extracted Requirements:")
    print(f"- Target: {requirements['target_columns']}")
    print(f"- Aggregation: {requirements['aggregation_functions']}")
    print(f"- Group By: {requirements['grouping_requirements']['group_by_columns']}")
    print()
    
    # Show how this would be used for formula generation
    print("Formula Generation Integration:")
    print("```python")
    print("# This would generate formulas like:")
    print("# =SUMIF(Sheet1!C:C,\"North\",Sheet1!B:B)  # North region")
    print("# =SUMIF(Sheet1!C:C,\"South\",Sheet1!B:B)  # South region")
    print("```")
    print()

if __name__ == "__main__":
    try:
        demo_formula_requirements_extraction()
        demo_integration_with_formula_generation()
        print("🎉 Demo completed successfully!")
    except Exception as e:
        print(f"❌ Demo failed: {str(e)}")
        import traceback
        traceback.print_exc()

