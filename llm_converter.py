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
Convert the given OCR text into a structured JSON format for restaurant menus.

CRITICAL INSTRUCTIONS:
1. Return ONLY valid JSON - no explanations, no markdown, no additional text
2. Follow the exact schema provided below
3. If information is missing, use the specified default values
4. Extract prices as numbers (remove currency symbols)
5. Identify menu categories and items accurately
6. Handle OCR errors gracefully by making reasonable inferences

JSON SCHEMA:
{
  "restaurant": {
    "name": "<restaurant name or 'Unknown'>",
    "source_image": "<path to the menu image or 'Unknown'>"
  },
  "categories": [
    {
      "categoryname": "<string>",
      "categoryid": null,
      "confidence": 1.0,
      "coordinates": null,
      "rank": null,
      "active": "1"
    }
  ],
  "items": [
    {
      "itemname": "<string>",
      "itemid": null,
      "categoryid": null,
      "description": "<string>",
      "price": <number or null>,
      "price_variants": [<array of numbers>],
      "currency": "INR",
      "instock": 2,
      "availability": 1,
      "tags": ["<infer tags like 'veg', 'non-veg', 'spicy', etc>"],
      "addongroups": [],
      "coordinates": null,
      "confidence": 1.0
    }
  ],
  "addongroups": [
    {
      "group_name": "<string>",
      "group_id": null,
      "min_select": 0,
      "max_select": 2,
      "items": ["<array of addon items>"]
    }
  ],
  "audit_log": []
}

DEFAULT VALUES:
- instock: 2
- availability: 1  
- active: "1"
- price: null (if not found)
- description: "" (if missing)
- currency: "INR"
- confidence: 1.0

EXAMPLES OF MENU ITEMS TO PARSE:
- "Chicken Teriyaki 259/-" → price: 259, itemname: "Chicken Teriyaki"
- "Add extra cheese 40/-" → addon item with price 40
- "CALZONE MENU" → category name
- "Japanese style with soy glaze" → item description

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
            json_data["restaurant"] = {"name": "Unknown", "source_image": source_image_path or "Unknown"}
        else:
            if source_image_path and json_data["restaurant"].get("source_image") in [None, "Unknown"]:
                json_data["restaurant"]["source_image"] = source_image_path
        
        if "categories" not in json_data:
            json_data["categories"] = []
        
        if "items" not in json_data:
            json_data["items"] = []
        
        if "addongroups" not in json_data:
            json_data["addongroups"] = []
        
        if "audit_log" not in json_data:
            json_data["audit_log"] = []
        
        # Validate and enhance categories
        for category in json_data["categories"]:
            category.setdefault("categoryid", None)
            category.setdefault("confidence", 1.0)
            category.setdefault("coordinates", None)
            category.setdefault("rank", None)
            category.setdefault("active", "1")
        
        # Validate and enhance items
        for item in json_data["items"]:
            item.setdefault("itemid", None)
            item.setdefault("categoryid", None)
            item.setdefault("description", "")
            item.setdefault("price_variants", [item.get("price")] if item.get("price") else [])
            item.setdefault("currency", "INR")
            item.setdefault("instock", 2)
            item.setdefault("availability", 1)
            item.setdefault("tags", [])
            item.setdefault("addongroups", [])
            item.setdefault("coordinates", None)
            item.setdefault("confidence", 1.0)
        
        # Validate and enhance addongroups
        for group in json_data["addongroups"]:
            group.setdefault("group_id", None)
            group.setdefault("min_select", 0)
            group.setdefault("max_select", 2)
            group.setdefault("items", [])
        
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