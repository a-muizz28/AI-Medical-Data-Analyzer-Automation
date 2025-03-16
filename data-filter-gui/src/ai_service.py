import json
import re
import google.generativeai as genai  # type: ignore

class AIService:
    def __init__(self, app):
        self.app = app
        
    def configure_api(self, api_key):
        """Configure the Gemini API with the provided key."""
        genai.configure(api_key=api_key)
        
    def extract_json(self, text):
        """Extract JSON data from text string."""
        try:
            # First attempt: Extract anything between square brackets
            match = re.search(r'\[.*\]', text, re.DOTALL)
            if match:
                json_text = match.group(0)
                # Validate by trying to parse it
                json.loads(json_text)
                return json_text
                
            # Second attempt: Look for any JSON-like structure
            pattern = r'(\{.*\}|\[.*\])'
            match = re.search(pattern, text, re.DOTALL)
            if match:
                json_text = match.group(0)
                # Validate it
                json.loads(json_text)
                return json_text
                
            return None
        except json.JSONDecodeError:
            # If we can't parse the extracted text as JSON, return None
            return None
            
    def ai_assisted_filter(self, df, filter_column, search_term):
        """
        Use AI to determine if rows match the search term semantically,
        not just by string containment.
        """
        self.app.add_to_status("Using AI to assist with filtering...")
        
        # Get a sample of the data to check
        sample_values = df[filter_column].astype(str).dropna().unique().tolist()
        if len(sample_values) > 20:  # Limit to 20 unique values for the prompt
            sample_values = sample_values[:20]
            
        # Create a prompt to check for semantic equivalence
        equivalence_prompt = f"""
        I'm looking for records related to "{search_term}" in a medical database.
        Below are some values from the {filter_column} column. 
        
        For each value, tell me if it is semantically equivalent to or related to "{search_term}".
        Answer in JSON format as follows:
        
        {{
            "matches": [
                "value1", "value2", ... (values that match)
            ],
            "explanation": "Brief explanation of the matches and equivalences"
        }}
        
        Values to check:
        {sample_values}
        """
        
        try:
            # Use a smaller, faster model for this filtering task
            ai_filter_model = genai.GenerativeModel("gemini-1.5-flash")
            filter_response = ai_filter_model.generate_content(equivalence_prompt)
            
            if filter_response.text:
                # Extract the JSON containing matched values
                match_json_text = self.extract_json(filter_response.text)
                if match_json_text:
                    matches_data = json.loads(match_json_text)
                    
                    # Get the list of matches
                    if "matches" in matches_data and isinstance(matches_data["matches"], list):
                        matched_values = matches_data["matches"]
                        
                        if matched_values:
                            self.app.add_to_status(f"AI found {len(matched_values)} related terms to '{search_term}'")
                            if "explanation" in matches_data:
                                self.app.add_to_status(f"AI explanation: {matches_data['explanation']}")
                            
                            # Create the filter based on these values
                            filter_mask = df[filter_column].astype(str).apply(
                                lambda x: any(match.lower() in x.lower() for match in matched_values)
                            )
                            
                            # Also include the original search term
                            filter_mask = filter_mask | df[filter_column].astype(str).str.contains(
                                search_term, case=False, na=False
                            )
                            
                            return df[filter_mask]
                        
            # Fallback to traditional filtering if AI doesn't provide useful results
            self.app.add_to_status("Falling back to standard filtering (AI didn't provide useful matches)")
            return df[df[filter_column].astype(str).str.contains(search_term, case=False, na=False)]
                    
        except Exception as filter_error:
            self.app.add_to_status(f"AI filtering error: {str(filter_error)}. Falling back to standard filtering.")
            return df[df[filter_column].astype(str).str.contains(search_term, case=False, na=False)]
            
    def analyze_data(self, search_term, filter_column, pdf_text, data_text):
        """Send data to AI for analysis and process the response."""
        # Prepare the AI prompt
        self.app.add_to_status("Preparing AI analysis...")
        
        prompt = f"""
        Analyze the following filtered data related to '{search_term}' in the {filter_column} column and provide insights based on the guidelines.

        {pdf_text}

        Filtered Data:
        {data_text}

        Provide the response in **JSON format** with the following structure:
        - For each record, include all the original data fields
        - Add TWO additional fields:
          1. "Meets Guidelines": MUST be one of exactly these three string values: 
             - "True" (fully or partially meets guidelines)
             - "False" (does not meet guidelines)
          2. "Notes on Compliance": A text explanation of your analysis.
        - Ensure patient/record identifiers match exactly with the original data
        - Accuracy and data integrity are crucial for the analysis.
        - Include any additional insights or recommendations based on the guidelines.
        - You are encouraged to provide detailed and informative responses.
        - You are a professional AI assistant specialized in medical data analysis.

        Example output format (with the actual columns from the data):

        [
            {{
                "column1": "value1",
                "column2": "value2",
                ...
                "Meets Guidelines": "True" or "False",
                "Notes on Compliance": "Treatment follows the guidelines for this condition."
            }},
            ...
        ]

        Ensure accuracy in extracting and formatting the response while maintaining data integrity.
        """

        # Send to Gemini AI
        self.app.add_to_status("Sending request to Gemini AI...")
        try:
            model = genai.GenerativeModel("gemini-1.5-flash") 
            response = model.generate_content(prompt)
            
            if not hasattr(response, 'text') or not response.text:
                raise ValueError("Empty response from AI")
                
            self.app.add_to_status("Processing AI response...")
            json_text = self.extract_json(response.text)
            if not json_text:
                raise ValueError("No valid JSON found in AI response.")

            response_json = json.loads(json_text)

            # Standardize boolean values
            for item in response_json:
                if "Meets Guidelines" in item:
                    if isinstance(item["Meets Guidelines"], str):
                        value = item["Meets Guidelines"].lower().strip()
                        # Set to True if exactly "true", otherwise False
                        item["Meets Guidelines"] = (value == "true")
                else:
                    # Default to False if missing
                    item["Meets Guidelines"] = False
                    
            return response_json
            
        except Exception as api_error:
            self.app.add_to_status(f"AI API Error: {str(api_error)}")
            return None
        