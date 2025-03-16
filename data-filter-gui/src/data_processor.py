import os
import pandas as pd  # type: ignore
import fitz  # type: ignore
from tkinter import messagebox, filedialog

class DataProcessor:
    def __init__(self, app):
        self.app = app
        
    def load_columns(self, excel_file_path, sheet_name):
        """Load columns from the specified Excel sheet."""
        if not excel_file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
            
        if not sheet_name:
            messagebox.showwarning("Warning", "Please enter a sheet name.")
            return
            
        try:
            self.app.add_to_status(f"Loading columns from sheet: {sheet_name}")
            
            # Load the Excel file and get the column names
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            self.app.available_columns = df.columns.tolist()
            
            # Update the dropdown with the column names
            self.app.filter_column_dropdown['values'] = self.app.available_columns
            
            if self.app.available_columns:
                # Select a default column (DiseaseName if available, otherwise first column)
                if 'DiseaseName' in self.app.available_columns:
                    self.app.filter_column.set('DiseaseName')
                else:
                    self.app.filter_column.set(self.app.available_columns[0])
                
                self.app.add_to_status(f"Loaded {len(self.app.available_columns)} columns from sheet: {sheet_name}")
            else:
                self.app.add_to_status("No columns found in the sheet.")
                
        except Exception as e:
            self.app.add_to_status(f"Error loading columns: {str(e)}")
            
            # Check if the sheet exists
            try:
                available_sheets = pd.ExcelFile(excel_file_path).sheet_names
                self.app.add_to_status(f"Available sheets: {', '.join(available_sheets)}")
            except:
                pass
                
    def process_data(self, filter_column, search_term, sheet_name, output_option):
        """Process the data with the given parameters."""
        self.app.add_to_status(f"Processing data where {filter_column} contains '{search_term}' in sheet: {sheet_name}")
        
        try:
            self.app.add_to_status("Reading Excel file...")
            try:
                df = pd.read_excel(self.app.excel_file_path, sheet_name=sheet_name)
            except ValueError as sheet_error:
                if "Worksheet named" in str(sheet_error) and "not found" in str(sheet_error):
                    available_sheets = pd.ExcelFile(self.app.excel_file_path).sheet_names
                    sheet_list = ", ".join(available_sheets)
                    self.app.add_to_status(f"Sheet '{sheet_name}' not found. Available sheets: {sheet_list}")
                    raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {sheet_list}")
                else:
                    raise sheet_error
                    
            if filter_column not in df.columns:
                raise ValueError(f"Column '{filter_column}' not found in the sheet.")

            # Filter the data based on the selected column and search term
            filtered_df = self.app.ai_service.ai_assisted_filter(df, filter_column, search_term)
            
            if filtered_df.empty:
                self.app.add_to_status(f"No data found where {filter_column} contains '{search_term}'.")
                return
                
            self.app.add_to_status(f"Found {len(filtered_df)} records where {filter_column} contains '{search_term}'")

            # Save the original data for later merging
            filtered_df = filtered_df.reset_index(drop=True)
            original_df = filtered_df.copy()

            # Read the PDF file
            self.app.add_to_status("Reading PDF file...")
            pdf_text = self._extract_pdf_text(self.app.pdf_file_path)
            self.app.add_to_status("PDF data extracted successfully")

            # Prepare data for AI
            data_text = filtered_df.to_string(index=False)

            # Generate AI response
            response_json = self.app.ai_service.analyze_data(
                search_term, filter_column, pdf_text, data_text
            )
            
            if not response_json:
                self.app.add_to_status("Failed to get analyzable response from AI.")
                return
                
            # Convert to DataFrame
            response_df = pd.DataFrame(response_json)
            
            # Save results
            self.app.file_utils.save_results(
                response_json, 
                response_df, 
                filter_column, 
                sheet_name, 
                output_option
            )
            
            self.app.add_to_status("Process completed successfully!")
                
        except Exception as e:
            self.app.add_to_status(f"Error: {str(e)}")
            
    def _extract_pdf_text(self, pdf_path):
        """Extract text from PDF file."""
        pdf_document = fitz.open(pdf_path)
        pdf_text = "\n".join([pdf_document.load_page(i).get_text() for i in range(pdf_document.page_count)])
        pdf_document.close()
        return pdf_text