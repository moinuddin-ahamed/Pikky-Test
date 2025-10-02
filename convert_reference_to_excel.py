"""
Convert data_reference.json to Excel format using the exporter mapper
"""

import json
import logging
import os
import sys
from pathlib import Path

# Add parent directory to path to import our modules
sys.path.insert(0, str(Path(__file__).parent))

from exporter import export_menu_to_excel

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s:%(name)s:%(message)s'
)

def load_json_file(file_path: str) -> dict:
    """Load JSON file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        logging.info(f"✅ Successfully loaded JSON from: {file_path}")
        return data
    except Exception as e:
        logging.error(f"❌ Failed to load JSON file: {e}")
        return None

def convert_reference_data_to_excel(input_json_path: str, output_dir: str):
    """
    Convert data_reference.json to Excel format
    
    Args:
        input_json_path: Path to data_reference.json file
        output_dir: Directory to save Excel output
    """
    print("🚀 Starting data_reference.json to Excel conversion...")
    print(f"📂 Input: {input_json_path}")
    print(f"📂 Output: {output_dir}")
    print()
    
    # Load JSON data
    data = load_json_file(input_json_path)
    if not data:
        print("❌ Failed to load JSON file")
        return False
    
    # Log structure summary
    print("📊 JSON Structure Summary:")
    print(f"  - Restaurants: {len(data.get('restaurants', []))}")
    print(f"  - Areas: {len(data.get('areas', []))}")
    print(f"  - Categories: {len(data.get('categories', []))}")
    print(f"  - Items: {len(data.get('items', []))}")
    print(f"  - Addon Groups: {len(data.get('addongroups', []))}")
    print()
    
    # The data_reference.json structure is already in the correct format
    # We just need to ensure it matches our expected structure
    
    # Transform to our expected structure
    transformed_data = {
        "restaurant": {
            "restaurantname": data.get('restaurants', [{}])[0].get('details', {}).get('restaurantname', 'Unknown'),
            "source_image": "data_reference.json",
            "country": data.get('restaurants', [{}])[0].get('details', {}).get('country', ''),
            "address": data.get('restaurants', [{}])[0].get('details', {}).get('address', ''),
            "contact": data.get('restaurants', [{}])[0].get('details', {}).get('contact', ''),
            "cuisines": data.get('restaurants', [{}])[0].get('details', {}).get('cuisines', ''),
            "city": data.get('restaurants', [{}])[0].get('details', {}).get('city', ''),
            "state": data.get('restaurants', [{}])[0].get('details', {}).get('state', ''),
        },
        "areas": data.get('areas', []),
        "categories": data.get('categories', []),
        "items": data.get('items', []),
        "addongroups": data.get('addongroups', []),
        "audit_log": []
    }
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Export to Excel
    try:
        excel_path, json_path = export_menu_to_excel(
            transformed_data,
            output_dir,
            "data_reference_converted",
            include_json=True,
            include_metadata=True,
            single_sheet=True
        )
        
        print("✅ Conversion successful!")
        print(f"📊 Excel file: {excel_path}")
        print(f"📄 JSON file: {json_path}")
        print()
        
        # Print some statistics
        print("📈 Export Statistics:")
        # Count total rows (approximate)
        total_items = len(transformed_data['items'])
        total_variations = sum(len(item.get('variation', [])) for item in transformed_data['items'])
        total_addon_refs = sum(len(item.get('addon', [])) for item in transformed_data['items'])
        
        print(f"  - Total Menu Items: {total_items}")
        print(f"  - Items with Variations: {total_variations}")
        print(f"  - Items with Addons: {total_addon_refs}")
        print()
        
        return True
        
    except Exception as e:
        logging.error(f"❌ Export failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main function"""
    # Default paths
    script_dir = Path(__file__).parent
    input_path = script_dir / "sample" / "data_reference.json"
    output_dir = script_dir / "output"
    
    # Check if input file exists
    if not input_path.exists():
        print(f"❌ Input file not found: {input_path}")
        return
    
    # Convert
    success = convert_reference_data_to_excel(str(input_path), str(output_dir))
    
    if success:
        print("🎉 All done! Check the output directory for your Excel file.")
    else:
        print("❌ Conversion failed. Check the logs above for details.")

if __name__ == "__main__":
    main()
