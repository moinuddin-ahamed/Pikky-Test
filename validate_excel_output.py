"""
Validate the converted Excel file structure
"""

import pandas as pd
import sys
from pathlib import Path

def validate_excel_file(excel_path: str):
    """Validate the Excel file has correct column headers and data"""
    print(f"🔍 Validating Excel file: {excel_path}")
    print()
    
    try:
        # Read the Excel file
        df = pd.read_excel(excel_path, sheet_name='Menu_Data')
        
        # Expected columns
        expected_columns = [
            "restaurant_name",
            "area_id",
            "area_display_name",
            "category_id",
            "category_name",
            "category_image_url",
            "category_timings",
            "category_rank",
            "item_id",
            "item_name",
            "item_description",
            "price",
            "rank",
            "image_url",
            "instock",
            "variation_item_id",
            "variation_id",
            "variation_name",
            "variation_price",
            "addon_name",
            "addon_item_selection",
            "addon_item_selection_min",
            "addon_item_selection_max",
            "addon_price",
            "addon_id",
            "addon_group_id",
            "addon_group_name",
        ]
        
        # Check columns
        actual_columns = df.columns.tolist()
        
        print("📊 Column Validation:")
        print(f"  Expected: {len(expected_columns)} columns")
        print(f"  Actual: {len(actual_columns)} columns")
        print()
        
        # Check if all expected columns are present
        missing_columns = [col for col in expected_columns if col not in actual_columns]
        extra_columns = [col for col in actual_columns if col not in expected_columns]
        
        if missing_columns:
            print("❌ Missing columns:")
            for col in missing_columns:
                print(f"  - {col}")
            print()
        
        if extra_columns:
            print("⚠️  Extra columns:")
            for col in extra_columns:
                print(f"  - {col}")
            print()
        
        if not missing_columns and not extra_columns:
            print("✅ All columns match perfectly!")
            print()
        
        # Display column headers
        print("📋 Column Headers:")
        for i, col in enumerate(actual_columns, 1):
            print(f"  {i:2d}. {col}")
        print()
        
        # Data statistics
        print("📈 Data Statistics:")
        print(f"  Total Rows: {len(df)}")
        print(f"  Unique Restaurants: {df['restaurant_name'].nunique()}")
        print(f"  Unique Categories: {df['category_name'].nunique()}")
        print(f"  Unique Items: {df['item_name'].nunique()}")
        print(f"  Items with Variations: {df['variation_name'].notna().sum()}")
        print(f"  Items with Addons: {df['addon_name'].notna().sum()}")
        print()
        
        # Sample data
        print("📝 Sample Data (first 3 rows):")
        print(df.head(3).to_string(index=False))
        print()
        
        # Check for data completeness
        print("🔎 Data Completeness:")
        critical_columns = ['restaurant_name', 'item_name', 'price']
        for col in critical_columns:
            null_count = df[col].isna().sum()
            null_pct = (null_count / len(df)) * 100
            status = "✅" if null_count == 0 else "⚠️"
            print(f"  {status} {col}: {null_count} nulls ({null_pct:.1f}%)")
        print()
        
        return True
        
    except Exception as e:
        print(f"❌ Validation failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main function"""
    # Find the most recent data_reference_converted Excel file
    output_dir = Path(__file__).parent / "output"
    
    excel_files = list(output_dir.glob("data_reference_converted_*.xlsx"))
    if not excel_files:
        print("❌ No converted Excel files found in output directory")
        return
    
    # Get the most recent file
    latest_file = max(excel_files, key=lambda p: p.stat().st_mtime)
    
    print(f"📂 Found Excel file: {latest_file.name}")
    print()
    
    # Validate
    success = validate_excel_file(str(latest_file))
    
    if success:
        print("🎉 Validation complete!")
    else:
        print("❌ Validation failed!")

if __name__ == "__main__":
    main()
