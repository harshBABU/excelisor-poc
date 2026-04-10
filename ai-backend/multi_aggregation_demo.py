#!/usr/bin/env python3
"""
Demo script for the Multi-Aggregation System.
Shows how to use the enhanced _calculate_multi_aggregation_grouped() method.
"""

import json
from typing import List, Dict, Any
from worksheet_manager import WorksheetManager

def create_sample_data() -> List[List[Any]]:
    """Create sample sales data for demonstration."""
    return [
        # Headers
        ["id", "sku_code", "region", "category", "qty", "revenue", "cost", "status", "date"],
        
        # Sample data rows
        ["A001", "SKU001", "North", "Electronics", 50, 1000, 600, "active", "2024-01-15"],
        ["A002", "SKU001", "North", "Electronics", 30, 800, 480, "active", "2024-01-16"],
        ["A003", "SKU001", "South", "Electronics", 40, 1200, 720, "active", "2024-01-17"],
        ["A004", "SKU002", "North", "Clothing", 60, 1500, 900, "active", "2024-01-18"],
        ["A005", "SKU002", "South", "Clothing", 35, 900, 540, "active", "2024-01-19"],
        ["A006", "SKU003", "East", "Home", 25, 750, 450, "inactive", "2024-01-20"],
        ["A007", "SKU001", "East", "Electronics", 45, 1100, 660, "active", "2024-01-21"],
        ["A008", "SKU002", "East", "Clothing", 55, 1300, 780, "active", "2024-01-22"],
        ["A009", "SKU004", "West", "Books", 80, 400, 240, "active", "2024-01-23"],
        ["A010", "SKU004", "West", "Books", 70, 350, 210, "active", "2024-01-24"],
        ["A011", "SKU001", "West", "Electronics", 65, 1400, 840, "active", "2024-01-25"],
        ["A012", "SKU005", "North", "Sports", 90, 1800, 1080, "active", "2024-01-26"],
        ["A013", "SKU005", "South", "Sports", 75, 1650, 990, "active", "2024-01-27"],
        ["A014", "SKU003", "North", "Home", 20, 600, 360, "active", "2024-01-28"],
        ["A015", "SKU006", "South", "Beauty", 40, 800, 480, "test", "2024-01-29"],
    ]

def demo_basic_multi_aggregation():
    """Demonstrate basic multi-aggregation functionality."""
    print("🔢 Basic Multi-Aggregation Demo")
    print("=" * 50)
    
    sample_data = create_sample_data()
    manager = WorksheetManager()
    
    # Example 1: Sum qty and average revenue by SKU and region
    print("\n📊 Example 1: Sum qty and average revenue by sku_code and region")
    print("-" * 60)
    
    pattern1 = {
        "aggregations": [
            {"column": "qty", "function": "SUM"},
            {"column": "revenue", "function": "AVERAGE"}
        ],
        "groupby_columns": ["sku_code", "region"],
        "filters": []
    }
    
    results1 = manager._calculate_multi_aggregation_grouped(
        sample_data, 
        pattern1,
        source_sheet="Sheet1"
    )
    
    print(f"✅ Found {results1['summary']['total_groups']} groups")
    print(f"📈 Processed {results1['summary']['aggregations_count']} aggregations")
    
    # Show first few results
    for i, result in enumerate(results1["results"][:3]):
        print(f"\nGroup {i+1}: {result['group_key']}")
        for agg_key, agg_data in result["aggregation_results"].items():
            print(f"  {agg_key}: {agg_data['value']} ({agg_data['function']} of {agg_data['column']})")
        print(f"  Formula example: {list(result['aggregation_results'].values())[0]['formula'][:60]}...")
    
    if len(results1["results"]) > 3:
        print(f"\n... and {len(results1['results']) - 3} more groups")

def demo_complex_multi_aggregation():
    """Demonstrate complex multi-aggregation with filters."""
    print("\n\n🎯 Complex Multi-Aggregation with Filters Demo")
    print("=" * 50)
    
    sample_data = create_sample_data()
    manager = WorksheetManager()
    
    # Example 2: Multiple aggregations with filters
    print("\n📊 Example 2: Count orders, sum revenue, average cost by category where status='active' and qty>30")
    print("-" * 90)
    
    pattern2 = {
        "aggregations": [
            {"column": "id", "function": "COUNT"},
            {"column": "revenue", "function": "SUM"},
            {"column": "cost", "function": "AVERAGE"},
            {"column": "qty", "function": "MAX"}
        ],
        "groupby_columns": ["category", "region"],
        "filters": [
            {"column": "status", "operator": "=", "value": "active", "logic": "AND"},
            {"column": "qty", "operator": ">", "value": "30", "logic": "AND"}
        ]
    }
    
    results2 = manager._calculate_multi_aggregation_grouped(
        sample_data, 
        pattern2,
        source_sheet="Sheet1"
    )
    
    print(f"✅ Found {results2['summary']['total_groups']} groups after filtering")
    print(f"📈 Processed {results2['summary']['aggregations_count']} aggregations")
    print(f"🔍 Applied {len(pattern2['filters'])} filters")
    print(f"📊 Filtered {results2['summary']['filtered_rows']} of {results2['summary']['original_rows']} rows")
    
    # Show detailed results
    for i, result in enumerate(results2["results"]):
        print(f"\nGroup {i+1}: {result['group_key']}")
        print(f"  Records: {result['row_count']}")
        for agg_key, agg_data in result["aggregation_results"].items():
            formatted_value = f"{agg_data['value']:,.2f}" if isinstance(agg_data['value'], float) else agg_data['value']
            print(f"  {agg_data['function']} of {agg_data['column']}: {formatted_value}")

def demo_natural_language_parsing():
    """Demonstrate natural language query parsing."""
    print("\n\n🗣️ Natural Language Query Parsing Demo")
    print("=" * 50)
    
    sample_data = create_sample_data()
    manager = WorksheetManager()
    available_columns = sample_data[0]  # Headers
    
    test_queries = [
        "sum qty and average revenue by sku_code and region",
        "count orders and total sales by category where status = 'active'",
        "max qty and min cost by region and category where revenue > 500",
        "average revenue by sku", # Should return None (not multi-agg)
        "sum qty, count id and average cost by category and region where status != 'test'"
    ]
    
    for i, query in enumerate(test_queries, 1):
        print(f"\n🔍 Query {i}: \"{query}\"")
        print("-" * (len(query) + 15))
        
        parsed_pattern = manager.parse_multi_aggregation_query(query, available_columns)
        
        if parsed_pattern:
            print("✅ Successfully parsed as multi-aggregation query")
            print(f"   Aggregations: {len(parsed_pattern['aggregations'])}")
            for agg in parsed_pattern['aggregations']:
                print(f"     - {agg['function']}({agg['column']})")
            
            print(f"   Group by: {', '.join(parsed_pattern['groupby_columns'])}")
            
            if parsed_pattern['filters']:
                print(f"   Filters: {len(parsed_pattern['filters'])}")
                for filter_cond in parsed_pattern['filters']:
                    print(f"     - {filter_cond['column']} {filter_cond['operator']} {filter_cond['value']}")
            
            # Execute the parsed query
            try:
                results = manager._calculate_multi_aggregation_grouped(
                    sample_data, 
                    parsed_pattern
                )
                print(f"   📊 Results: {results['summary']['total_groups']} groups")
            except Exception as e:
                print(f"   ❌ Execution failed: {str(e)}")
        else:
            print("❌ Not detected as multi-aggregation query (this is normal for simple queries)")

def demo_worksheet_creation():
    """Demonstrate worksheet creation from multi-aggregation results."""
    print("\n\n📋 Worksheet Creation Demo")
    print("=" * 50)
    
    sample_data = create_sample_data()
    manager = WorksheetManager()
    
    # Create a comprehensive analysis
    pattern = {
        "aggregations": [
            {"column": "qty", "function": "SUM"},
            {"column": "revenue", "function": "AVERAGE"},
            {"column": "id", "function": "COUNT"}
        ],
        "groupby_columns": ["category", "region"],
        "filters": [
            {"column": "status", "operator": "=", "value": "active", "logic": "AND"}
        ]
    }
    
    results = manager._calculate_multi_aggregation_grouped(sample_data, pattern)
    worksheet = manager.create_multi_aggregation_worksheet(results, "Sales Analysis")
    
    print(f"✅ Created worksheet: {worksheet['name']}")
    print(f"📊 Cells: {len(worksheet['cells'])}")
    print(f"🧮 Formulas: {len(worksheet['formulas'])}")
    print(f"📝 Notes: {worksheet.get('notes', 'None')}")
    
    # Show sample cell content
    print("\n📋 Sample Worksheet Content:")
    print("-" * 30)
    
    # Group cells by row for display
    rows = {}
    for cell in worksheet['cells'][:20]:  # Show first 20 cells
        address = cell['address']
        row_num = ''.join(filter(str.isdigit, address))
        if row_num not in rows:
            rows[row_num] = []
        rows[row_num].append(cell)
    
    for row_num in sorted(rows.keys(), key=int):
        row_cells = sorted(rows[row_num], key=lambda x: x['address'])
        values = [str(cell['value'])[:15] for cell in row_cells]
        print(f"Row {row_num}: {' | '.join(values)}")
    
    if len(worksheet['cells']) > 20:
        print(f"... and {len(worksheet['cells']) - 20} more cells")

def demo_performance_features():
    """Demonstrate performance features with larger datasets."""
    print("\n\n⚡ Performance Features Demo")
    print("=" * 50)
    
    print("🔧 Creating larger synthetic dataset...")
    
    # Create larger dataset for performance testing
    large_data = [["id", "sku", "region", "sales", "qty"]]  # Headers
    
    skus = [f"SKU{i:03d}" for i in range(1, 21)]  # 20 SKUs
    regions = ["North", "South", "East", "West", "Central"]
    
    # Generate 1000 records
    for i in range(1, 1001):
        sku = skus[(i-1) % len(skus)]
        region = regions[(i-1) % len(regions)]
        sales = 100 + (i % 500) * 10  # Varying sales amounts
        qty = 1 + (i % 50)  # Varying quantities
        large_data.append([f"ID{i:04d}", sku, region, sales, qty])
    
    print(f"✅ Created dataset with {len(large_data)-1} records")
    
    manager = WorksheetManager()
    
    # Performance test with complex aggregation
    import time
    start_time = time.time()
    
    large_pattern = {
        "aggregations": [
            {"column": "sales", "function": "SUM"},
            {"column": "sales", "function": "AVERAGE"},
            {"column": "qty", "function": "SUM"},
            {"column": "id", "function": "COUNT"}
        ],
        "groupby_columns": ["sku", "region"],
        "filters": [
            {"column": "sales", "operator": ">", "value": "200", "logic": "AND"}
        ]
    }
    
    results = manager._calculate_multi_aggregation_grouped(large_data, large_pattern)
    
    end_time = time.time()
    processing_time = end_time - start_time
    
    print(f"⏱️  Processing time: {processing_time:.3f} seconds")
    print(f"📊 Groups generated: {results['summary']['total_groups']}")
    print(f"🔍 Rows after filtering: {results['summary']['filtered_rows']} of {results['summary']['original_rows']}")
    print(f"📈 Aggregations calculated: {results['summary']['aggregations_count']}")
    print(f"⚡ Processing rate: {int(results['summary']['filtered_rows'] / processing_time):,} rows/second")
    
    # Memory efficiency
    import sys
    result_size = sys.getsizeof(results)
    print(f"💾 Result size: {result_size:,} bytes ({result_size/1024:.1f} KB)")

def main():
    """Run all demo functions."""
    print("🚀 Multi-Aggregation System Demo")
    print("=" * 60)
    print("This demo showcases the enhanced multi-aggregation capabilities")
    print("for handling complex Excel worksheet generation scenarios.")
    print()
    
    try:
        demo_basic_multi_aggregation()
        demo_complex_multi_aggregation()
        demo_natural_language_parsing()
        demo_worksheet_creation()
        demo_performance_features()
        
        print("\n\n🎉 All demos completed successfully!")
        print("\nThe multi-aggregation system is ready for production use.")
        print("\nKey benefits demonstrated:")
        print("✅ Multiple aggregations in single query")
        print("✅ Multiple grouping columns support")
        print("✅ Advanced filtering capabilities")
        print("✅ Natural language query parsing")
        print("✅ Professional worksheet generation")
        print("✅ High performance with large datasets")
        
    except Exception as e:
        print(f"❌ Demo failed with error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

