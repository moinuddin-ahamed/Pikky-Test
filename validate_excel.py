"""
Quick validation script to check Excel output structure
"""
import pandas as pd
import os

def validate_excel_structure(excel_path):
    """Validate that the Excel file has the expected column headers"""
    
    expected_columns = [
        "restaurant_name", "area_id", "area_display_name", "category_id", 
        "category_name", "category_image_url", "category_timings", "category_rank",
        "item_id", "item_name", "item_description", "price", "rank", 
        "image_url", "instock", "variation_item_id", "variation_id", 
        "variation_name", "variation_price", "addon_name", "addon_item_selection",
        "addon_item_selection_min", "addon_item_selection_max", "addon_price",
        "addon_id", "addon_group_id", "addon_group_name"
    ]
    
    try:
        # Read the Excel file
        df = pd.read_excel(excel_path, sheet_name="Menu_Data")
        
        print(f"✅ Successfully loaded Excel file: {excel_path}")
        print(f"📊 Total rows: {len(df)}")
        print(f"📊 Total columns: {len(df.columns)}")
        
        # Check columns
        actual_columns = list(df.columns)
        print(f"\n🔍 Column validation:")
        
        missing_columns = []
        extra_columns = []
        
        for col in expected_columns:
            if col in actual_columns:
                print(f"  ✅ {col}")
            else:
                print(f"  ❌ {col} - MISSING")
                missing_columns.append(col)
        
        for col in actual_columns:
            if col not in expected_columns:
                print(f"  ⚠️ {col} - EXTRA")
                extra_columns.append(col)
        
        # Show sample data
        print(f"\n📋 Sample data (first 3 rows):")
        print(df.head(3).to_string())
        
        # Summary
        print(f"\n📊 Validation Summary:")
        print(f"Expected columns: {len(expected_columns)}")
        print(f"Actual columns: {len(actual_columns)}")
        print(f"Missing columns: {len(missing_columns)}")
        print(f"Extra columns: {len(extra_columns)}")
        
        if missing_columns:
            print(f"Missing: {missing_columns}")
        if extra_columns:
            print(f"Extra: {extra_columns}")
        
        return len(missing_columns) == 0 and len(extra_columns) == 0
        
    except Exception as e:
        print(f"❌ Error validating Excel file: {e}")
        return False

if __name__ == "__main__":
    # Find the latest Excel file
    output_dir = "output"
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    
    if excel_files:
        latest_file = max(excel_files)
        excel_path = os.path.join(output_dir, latest_file)
        print(f"🔍 Validating: {latest_file}")
        
        success = validate_excel_structure(excel_path)
        if success:
            print("\n🎉 Excel structure validation PASSED!")
        else:
            print("\n⚠️ Excel structure validation FAILED!")
    else:
        print("❌ No Excel files found in output directory")