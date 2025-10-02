"""
LLM Converter Module for converting OCR text to structured JSON using Google Gemini
"""

import json
import logging
import os
from typing import Dict, Optional, Tuple

try:
    import google.generativeai as genai
except ImportError:
    genai = None
    logging.warning("google-generativeai package not found. Please install it to use LLM features.")


class GeminiConverter:
    """
    Converter class for processing OCR text using Google Gemini LLM
    """
    
    def __init__(self, model_name: str = "gemini-2.0-flash-exp"):
        """
        Initialize the Gemini converter
        
        Args:
            model_name: The Gemini model to use (gemini-2.0-flash-exp, gemini-2.5-pro, etc.)
        """
        self.model_name = model_name
        self.model = None
        self._initialize_model()
    
    def _initialize_model(self) -> bool:
        """
        Initialize the Gemini model with API key
        
        Returns:
            bool: True if successful, False otherwise
        """
        if genai is None:
            logging.error("google-generativeai package not installed")
            return False
        
        # Get API key from environment variable
        api_key = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
        if not api_key:
            logging.error(
                "Gemini API key not found. Please set GOOGLE_API_KEY or GEMINI_API_KEY environment variable"
            )
            return False
        
        try:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel(self.model_name)
            logging.info(f"Successfully initialized Gemini model: {self.model_name}")
            return True
        except Exception as e:
            logging.error(f"Failed to initialize Gemini model: {e}")
            return False
    
    def get_system_prompt(self) -> str:
        """
        Get the system prompt for menu parsing
        
        Returns:
            str: The system prompt for Gemini
        """
        return """You are an expert at parsing restaurant menu text extracted from images using OCR. 
Convert the given OCR text into a structured JSON format matching a restaurant POS system database schema.

CRITICAL INSTRUCTIONS:
1. Return ONLY valid JSON - no explanations, no markdown, no additional text
2. Follow the exact schema provided below
3. If information is missing, use the specified default values
4. Extract prices as numbers (remove currency symbols like ₹, -, /)
5. Identify menu categories and items accurately
6. Handle OCR errors gracefully by making reasonable inferences
7. For items with multiple prices (e.g., "289/349"), create variation records

JSON SCHEMA:
{
  "restaurant": {
    "restaurantname": "<restaurant name or 'Unknown'>",
    "source_image": "<path to the menu image>",
    "country": "",
    "address": "",
    "contact": "",
    "cuisines": "",
    "city": "",
    "state": ""
  },
  "areas": [
    {
      "areaid": null,
      "displayname": "Main Dining",
      "active": "1",
      "rank": "1"
    }
  ],
  "categories": [
    {
      "categoryid": null,
      "categoryname": "<string>",
      "active": "1",
      "categoryrank": "<sequence number>",
      "category_image_url": null,
      "parent_category_id": "0",
      "categorytimings": ""
    }
  ],
  "items": [
    {
      "itemid": null,
      "itemallowvariation": 0,
      "itemname": "<string>",
      "itemrank": "1",
      "item_categoryid": null,
      "price": "<price as string>",
      "active": "1",
      "item_favorite": "0",
      "itemallowaddon": "1",
      "itemaddonbasedon": "0",
      "instock": "2",
      "ignore_taxes": "0",
      "ignore_discounts": "0",
      "days": "-1",
      "item_attributeid": "<1=veg, 2=non-veg, 24=egg>",
      "itemdescription": "<string>",
      "minimumpreparationtime": "",
      "item_image_url": "",
      "variation": [
        {
          "variationitemid": null,
          "variationid": null,
          "itemid": null,
          "variation_name": "<size/type>",
          "variation_price": "<price as string>",
          "variationrank": "1"
        }
      ],
      "addon": [
        {
          "addon_group_id": null,
          "addon_item_selection": "M",
          "addon_item_selection_min": "0",
          "addon_item_selection_max": "2"
        }
      ],
      "item_tax": ""
    }
  ],
  "addongroups": [
    {
      "addongroupid": null,
      "addongroup_restaurantid": null,
      "addongroup_rank": "1",
      "active": "1",
      "show_in_online": "1",
      "show_in_pos": "1",
      "min_qty": "0",
      "max_qty": "2",
      "addongroup_name": "<string>",
      "addongroupitems": [
        {
          "addonitemid": null,
          "addonitem_name": "<string>",
          "addonitem_price": "<price as string>",
          "active": "1",
          "attributes": "<1=veg, 2=non-veg>",
          "addonitem_rank": "1",
          "parent_addon_id": "0",
          "status": "1"
        }
      ]
    }
  ],
  "audit_log": []
}

DEFAULT VALUES:
- instock: "2" (2=available, 1=low stock, 0=out of stock)
- active: "1" (1=active, 0=inactive)
- item_attributeid: "1" for veg, "2" for non-veg, "24" for egg
- itemallowaddon: "1" if item can have add-ons
- itemallowvariation: 0 (set to 1 if item has size/type variations)
- categoryrank: sequential number based on order in menu
- itemrank: "1" by default

PARSING RULES:
1. Category Detection: Lines with ALL CAPS or section headers (e.g., "CALZONE MENU", "PIZZAS")
2. Item Format: "Item Name   Price" or "Item Name   ₹Price/-"
3. Description: Text in parentheses or lines following item name
4. Price Variants: "289/349" → create variation entries
5. Add-ons: Items like "Add extra cheese 40" → goes to addongroups
6. Attributes: Detect veg/non-veg from context (chicken/mutton=non-veg, paneer/veg=veg)

EXAMPLES:
Input: "Three Cheese ₹259/-\n(Mozzarella+Cheddar+Cream Cheese)"
Output: itemname="Three Cheese", price="259", itemdescription="(Mozzarella+Cheddar+Cream Cheese)", item_attributeid="1"

Input: "Chicken Teriyaki 289/349"
Output: itemname="Chicken Teriyaki", price="289", item_attributeid="2", itemallowvariation=1, 
        variation=[{variation_name: "Regular", variation_price: "289"}, {variation_name: "Large", variation_price: "349"}]

Input: "Add extra cheese 40/-"
Output: addongroup with addonitem_name="Extra Cheese", addonitem_price="40"

Return ONLY the JSON object, nothing else."""

    def text_to_json_with_gemini(
        self, 
        raw_text: str, 
        source_image_path: Optional[str] = None
    ) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Convert OCR text to structured JSON using Gemini LLM
        
        Args:
            raw_text: The OCR extracted text from menu image
            source_image_path: Optional path to the source image
        
        Returns:
            Tuple of (success, json_data, error_message)
        """
        if not self.model:
            return False, None, "Gemini model not initialized"
        
        if not raw_text or not raw_text.strip():
            return False, None, "Empty or invalid input text"
        
        try:
            # Prepare the prompt
            system_prompt = self.get_system_prompt()
            user_prompt = f"""OCR Text from Menu Image:
{raw_text.strip()}

Source Image: {source_image_path or 'Unknown'}

Convert this to the specified JSON format:"""
            
            full_prompt = f"{system_prompt}\n\n{user_prompt}"
            
            # Generate response from Gemini
            logging.info("Sending request to Gemini...")
            response = self.model.generate_content(full_prompt)
            
            if not response or not response.text:
                return False, None, "Empty response from Gemini"
            
            # Clean the response text
            response_text = response.text.strip()
            
            # Remove any markdown code blocks if present
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
            
            response_text = response_text.strip()
            
            # Parse JSON
            try:
                json_data = json.loads(response_text)
                
                # Validate and enhance the JSON structure
                json_data = self._validate_and_enhance_json(json_data, source_image_path)
                
                logging.info("Successfully converted text to JSON using Gemini")
                return True, json_data, None
                
            except json.JSONDecodeError as e:
                logging.error(f"Failed to parse JSON from Gemini response: {e}")
                logging.debug(f"Raw response: {response_text[:500]}...")
                return False, None, f"Invalid JSON format from Gemini: {e}"
        
        except Exception as e:
            logging.error(f"Error calling Gemini API: {e}")
            return False, None, f"Gemini API error: {e}"
    
    def _validate_and_enhance_json(
        self, 
        json_data: Dict, 
        source_image_path: Optional[str] = None
    ) -> Dict:
        """
        Validate and enhance the JSON structure with defaults and corrections
        
        Args:
            json_data: The parsed JSON data from Gemini
            source_image_path: Optional path to source image
        
        Returns:
            Dict: Enhanced and validated JSON data
        """
        # Ensure top-level structure
        if "restaurant" not in json_data:
            json_data["restaurant"] = {
                "restaurantname": "Unknown",
                "source_image": source_image_path or "Unknown",
                "country": "",
                "address": "",
                "contact": "",
                "cuisines": "",
                "city": "",
                "state": ""
            }
        else:
            if source_image_path:
                json_data["restaurant"]["source_image"] = source_image_path
            json_data["restaurant"].setdefault("restaurantname", "Unknown")
            json_data["restaurant"].setdefault("country", "")
            json_data["restaurant"].setdefault("address", "")
            json_data["restaurant"].setdefault("contact", "")
            json_data["restaurant"].setdefault("cuisines", "")
            json_data["restaurant"].setdefault("city", "")
            json_data["restaurant"].setdefault("state", "")
        
        if "areas" not in json_data:
            json_data["areas"] = []
        
        if "categories" not in json_data:
            json_data["categories"] = []
        
        if "items" not in json_data:
            json_data["items"] = []
        
        if "addongroups" not in json_data:
            json_data["addongroups"] = []
        
        if "audit_log" not in json_data:
            json_data["audit_log"] = []
        
        # Validate and enhance areas
        for area in json_data["areas"]:
            area.setdefault("areaid", None)
            area.setdefault("displayname", "Main Dining")
            area.setdefault("active", "1")
            area.setdefault("rank", "1")
        
        # Validate and enhance categories
        for category in json_data["categories"]:
            category.setdefault("categoryid", None)
            category.setdefault("active", "1")
            category.setdefault("categoryrank", "1")
            category.setdefault("category_image_url", None)
            category.setdefault("parent_category_id", "0")
            category.setdefault("categorytimings", "")
        
        # Validate and enhance items
        for item in json_data["items"]:
            item.setdefault("itemid", None)
            item.setdefault("itemallowvariation", 0)
            item.setdefault("itemrank", "1")
            item.setdefault("item_categoryid", None)
            item.setdefault("active", "1")
            item.setdefault("item_favorite", "0")
            item.setdefault("itemallowaddon", "1")
            item.setdefault("itemaddonbasedon", "0")
            item.setdefault("instock", "2")
            item.setdefault("ignore_taxes", "0")
            item.setdefault("ignore_discounts", "0")
            item.setdefault("days", "-1")
            item.setdefault("item_attributeid", "1")
            item.setdefault("itemdescription", "")
            item.setdefault("minimumpreparationtime", "")
            item.setdefault("item_image_url", "")
            item.setdefault("variation", [])
            item.setdefault("addon", [])
            item.setdefault("item_tax", "")
        
        # Validate and enhance addongroups
        for group in json_data["addongroups"]:
            group.setdefault("addongroupid", None)
            group.setdefault("addongroup_restaurantid", None)
            group.setdefault("addongroup_rank", "1")
            group.setdefault("active", "1")
            group.setdefault("show_in_online", "1")
            group.setdefault("show_in_pos", "1")
            group.setdefault("min_qty", "0")
            group.setdefault("max_qty", "2")
            group.setdefault("addongroupitems", [])
            
            # Validate addon group items
            for addon_item in group.get("addongroupitems", []):
                addon_item.setdefault("addonitemid", None)
                addon_item.setdefault("active", "1")
                addon_item.setdefault("attributes", "1")
                addon_item.setdefault("addonitem_rank", "1")
                addon_item.setdefault("parent_addon_id", "0")
                addon_item.setdefault("status", "1")
        
        return json_data


def text_to_json_with_gemini(
    raw_text: str, 
    source_image_path: Optional[str] = None,
    model_name: str = "gemini-2.0-flash-exp"
) -> Tuple[bool, Optional[Dict], Optional[str]]:
    """
    Convenience function to convert text to JSON using Gemini
    
    Args:
        raw_text: The OCR extracted text
        source_image_path: Optional path to source image
        model_name: Gemini model to use
    
    Returns:
        Tuple of (success, json_data, error_message)
    """
    converter = GeminiConverter(model_name)
    return converter.text_to_json_with_gemini(raw_text, source_image_path)


# Example usage and testing
if __name__ == "__main__":
    # Test with sample menu text
    sample_text = """CALZONE MENU
Chicken Teriyaki 259/-
Japanese style with soy glaze
Add extra cheese 40/-
Vegetable Supreme 229/-
Fresh vegetables with herbs
Extra toppings 30/-"""
    
    success, json_data, error = text_to_json_with_gemini(sample_text, "sample/menu.jpg")
    
    if success:
        print("✅ Conversion successful!")
        print(json.dumps(json_data, indent=2))
    else:
        print(f"❌ Conversion failed: {error}")