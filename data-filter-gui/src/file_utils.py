"""
File utilities for the AI Medical Data Analyzer Application.
Handles file operations like reading Excel/PDF files and saving output files.
"""

import os
import json
import pandas as pd  # type: ignore
import fitz  # type: ignore
from tkinter import filedialog, messagebox


def read_excel_file(file_path, sheet_name):
    """
    Read data from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to read
        
    Returns:
        tuple: (dataframe or None, error message or None)
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df, None
    except ValueError as sheet_error:
        if "Worksheet named" in str(sheet_error) and "not found" in str(sheet_error):
            available_sheets = pd.ExcelFile(file_path).sheet_names
            sheet_list = ", ".join(available_sheets)
            error_msg = f"Sheet '{sheet_name}' not found. Available sheets: {sheet_list}"
            return None, error_msg
        else:
            return None, str(sheet_error)
    except Exception as e:
        return None, str(e)


def get_available_sheets(excel_file_path):
    """
    Get list of available sheets in an Excel file.
    
    Args:
        excel_file_path: Path to the Excel file
        
    Returns:
        list: List of sheet names, or empty list if error
    """
    try:
        return pd.ExcelFile(excel_file_path).sheet_names
    except Exception:
        return []


def read_pdf_file(pdf_file_path):
    """
    Extract text from a PDF file.
    
    Args:
        pdf_file_path: Path to the PDF file
        
    Returns:
        tuple: (pdf_text or None, error message or None)
    """
    try:
        pdf_document = fitz.open(pdf_file_path)
        pdf_text = "\n".join([pdf_document.load_page(i).get_text() for i in range(pdf_document.page_count)])
        pdf_document.close()
        return pdf_text, None
    except Exception as e:
        return None, str(e)


def save_to_json(response_json, excel_file_path, filter_column, is_new_file=True):
    """
    Save data to a JSON file, either new or appending to existing.
    
    Args:
        response_json: JSON data to save
        excel_file_path: Path to the source Excel file (for naming)
        filter_column: Column used for filtering (for naming)
        is_new_file: Whether to create a new file or append to existing
        
    Returns:
        tuple: (output_path or None, error message or None)
    """
    base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    default_json_name = f"{base_name}_by_{filter_column}_Analyzed.json"
    
    if is_new_file:
        # Create a new JSON file
        output_json_path = filedialog.asksaveasfilename(
            defaultextension=".json", 
            filetypes=[("JSON files", "*.json")],
            title="Save JSON results",
            initialfile=default_json_name
        )
        
        if not output_json_path:
            return None, "JSON file save cancelled."
            
        try:
            with open(output_json_path, "w", encoding="utf-8") as json_file:
                json.dump(response_json, json_file, indent=4)
            return output_json_path, None
        except Exception as e:
            return None, f"Error saving JSON: {str(e)}"
    else:
        # Add to existing JSON file
        target_json_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            title="Select JSON file to append data to (or Cancel for new file)"
        )
        
        if not target_json_path:
            # User cancelled, ask if they want to create a new JSON file
            if messagebox.askyesno("JSON File Selection", 
                                "No existing JSON file selected. Would you like to save to a new JSON file?"):
                return save_to_json(response_json, excel_file_path, filter_column, True)
            else:
                return None, "JSON save cancelled."
        
        try:
            with open(target_json_path, "r", encoding="utf-8") as json_file:
                try:
                    existing_data = json.load(json_file)
                    
                    # Check if the existing data is a list
                    if isinstance(existing_data, list):
                        # Append the new data to the existing list
                        combined_data = existing_data + response_json
                        
                        # Write back the combined data
                        with open(target_json_path, "w", encoding="utf-8") as out_file:
                            json.dump(combined_data, out_file, indent=4)
                        
                        return target_json_path, None
                    else:
                        return None, "Existing JSON file is not in the expected list format"
                        
                except json.JSONDecodeError:
                    return None, "Selected file does not contain valid JSON data"
        except Exception as json_error:
            # Ask if user wants to create a new JSON file instead
            if messagebox.askyesno("JSON Error", 
                                "Could not append to the selected JSON file. Would you like to save to a new file instead?"):
                return save_to_json(response_json, excel_file_path, filter_column, True)
            else:
                return None, f"Error appending to JSON: {str(json_error)}"


def save_to_excel(response_df, excel_file_path, sheet_name, filter_column, is_new_file=True):
    """
    Save dataframe to Excel file, either new or as a new sheet in existing file.
    
    Args:
        response_df: DataFrame to save
        excel_file_path: Path to the source Excel file (for naming)
        sheet_name: Original sheet name (for naming new sheet)
        filter_column: Column used for filtering (for naming)
        is_new_file: Whether to create a new file or add sheet to existing
        
    Returns:
        tuple: (output_path or None, error message or None)
    """
    base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    analyzed_sheet_name = f"{sheet_name}_by_{filter_column}_Analyzed"
    
    if is_new_file:
        # Create a new Excel file
        default_excel_name = f"{base_name}_by_{filter_column}_Analyzed.xlsx"
        
        output_excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")],
            title="Save as new Excel file",
            initialfile=default_excel_name
        )
        
        if not output_excel_path:
            return None, "Excel file save cancelled."
            
        try:
            response_df.to_excel(output_excel_path, sheet_name=analyzed_sheet_name, index=False)
            return output_excel_path, None
        except Exception as e:
            return None, f"Error saving Excel file: {str(e)}"
    else:
        # Add as new sheet to existing file
        target_excel_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Select Excel file to add sheet to"
        )
        
        if not target_excel_path:
            return None, "No target Excel file selected. Operation cancelled."
            
        try:
            # Check if the sheet already exists
            with pd.ExcelFile(target_excel_path) as xls:
                existing_sheets = xls.sheet_names
                
            counter = 1
            original_name = analyzed_sheet_name
            while analyzed_sheet_name in existing_sheets:
                analyzed_sheet_name = f"{original_name}_{counter}"
                counter += 1
                
            # Write to the file
            try:
                with pd.ExcelWriter(target_excel_path, engine='openpyxl', mode='a') as writer:
                    response_df.to_excel(writer, sheet_name=analyzed_sheet_name, index=False)
                
                return target_excel_path, None
            except PermissionError:
                error_msg = "Cannot write to the Excel file. It may be open in another program."
                messagebox.showerror("Permission Error", error_msg)
                return None, error_msg
            except Exception as write_error:
                # Ask user if they want to save to a new file instead
                if messagebox.askyesno("Excel Write Error", 
                                      "Could not write to the selected Excel file. Would you like to save to a new file instead?"):
                    default_excel_name = f"{base_name}_by_{filter_column}_Analyzed.xlsx"
                    
                    new_excel_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx", 
                        filetypes=[("Excel files", "*.xlsx")],
                        title="Save as new Excel file",
                        initialfile=default_excel_name
                    )
                    
                    if new_excel_path:
                        response_df.to_excel(new_excel_path, sheet_name=analyzed_sheet_name, index=False)
                        return new_excel_path, None
                    else:
                        return None, "Excel file save cancelled."
                else:
                    return None, f"Error writing to Excel file: {str(write_error)}"
        except Exception as e:
            # Ask if they want to try saving to a new file instead
            if messagebox.askyesno("Excel Error", 
                                  "Error working with the selected Excel file. Would you like to save to a new file instead?"):
                default_excel_name = f"{base_name}_by_{filter_column}_Analyzed.xlsx"
                
                new_excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", 
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save as new Excel file",
                    initialfile=default_excel_name
                )
                
                if new_excel_path:
                    response_df.to_excel(new_excel_path, sheet_name=analyzed_sheet_name, index=False)
                    return new_excel_path, None
                else:
                    return None, "Excel file save cancelled."
            else:
                return None, f"Error checking existing sheets: {str(e)}"


def extract_json_from_text(text):
    """
    Extract JSON from AI response text.
    
    Args:
        text: Text to extract JSON from
        
    Returns:
        str: Extracted JSON text or None if not found
    """
    import re
    
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