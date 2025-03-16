import json
import pandas as pd #type: ignore
import google.generativeai as genai  #type: ignore
import fitz  #type: ignore
import os 
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from ttkthemes import ThemedTk #type: ignore
from dotenv import load_dotenv #type: ignore

class DataFilterApp:
    def __init__(self, master):
        self.master = master
        master.title("AI Data Analyzer Application")
        
        # Set window size and position
        window_width = 750
        window_height = 800
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x_coordinate = int((screen_width/2) - (window_width/2))
        y_coordinate = int((screen_height/2) - (window_height/2))
        master.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

        # Initialize file paths
        self.excel_file_path = ""
        self.pdf_file_path = ""
        self.output_json_path = ""
        self.output_excel_path = ""
        
        # Available columns in the loaded Excel file
        self.available_columns = []
        
        # Output options
        self.output_option = tk.StringVar(value="new_file")

        # Load environment variables from .env file (if it exists)
        load_dotenv()

        # Get API key from environment variables with fallback
        self.gemini_api_key = os.environ.get("GEMINI_API_KEY")
        if not self.gemini_api_key:
            messagebox.showerror("API Key Error", "Gemini API key not found in environment variables. Please set GEMINI_API_KEY.")
        else:
            genai.configure(api_key=self.gemini_api_key)
        

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", font=("Helvetica", 10))
        self.style.configure("TLabel", font=("Helvetica", 11))
        self.style.configure("Header.TLabel", font=("Helvetica", 16, "bold"))
        self.style.configure("Section.TFrame", padding=10)

        self.create_widgets()

    def create_widgets(self):
        # Main container with padding
        main_frame = ttk.Frame(self.master, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        header_label = ttk.Label(header_frame, text="AI Medical Data Analyzer Application", style="Header.TLabel")
        header_label.pack()
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="Input Files", padding=15)
        files_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Excel file selection
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_status = tk.StringVar(value="No file selected")
        excel_label = ttk.Label(excel_frame, text="Excel File:")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        
        excel_status_label = ttk.Label(excel_frame, textvariable=self.excel_status, foreground="gray")
        excel_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        select_excel_button = ttk.Button(excel_frame, text="Browse...", command=self.select_excel_file)
        select_excel_button.pack(side=tk.RIGHT)
        
        # PDF file selection
        pdf_frame = ttk.Frame(files_frame)
        pdf_frame.pack(fill=tk.X, pady=5)
        
        self.pdf_status = tk.StringVar(value="No file selected")
        pdf_label = ttk.Label(pdf_frame, text="PDF File:")
        pdf_label.pack(side=tk.LEFT, padx=(0, 5))
        
        pdf_status_label = ttk.Label(pdf_frame, textvariable=self.pdf_status, foreground="gray")
        pdf_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        select_pdf_button = ttk.Button(pdf_frame, text="Browse...", command=self.select_pdf_file)
        select_pdf_button.pack(side=tk.RIGHT)
        
        # Parameters section
        params_frame = ttk.LabelFrame(main_frame, text="Filter Parameters", padding=15)
        params_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Sheet name entry
        sheet_frame = ttk.Frame(params_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        sheet_label = ttk.Label(sheet_frame, text="Excel Sheet Name:")
        sheet_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.sheet_name_entry = ttk.Entry(sheet_frame, width=30)
        self.sheet_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.sheet_name_entry.insert(0, "Sheet1")
        
        refresh_columns_button = ttk.Button(sheet_frame, text="Load Columns", command=self.load_columns)
        refresh_columns_button.pack(side=tk.RIGHT)
        
        # Filter column selection
        filter_col_frame = ttk.Frame(params_frame)
        filter_col_frame.pack(fill=tk.X, pady=5)
        
        filter_col_label = ttk.Label(filter_col_frame, text="Filter Column:")
        filter_col_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.filter_column = tk.StringVar()
        self.filter_column_dropdown = ttk.Combobox(filter_col_frame, textvariable=self.filter_column, state="readonly", width=28)
        self.filter_column_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.filter_column_dropdown['values'] = ["<Load Excel file first>"]
        
        # Search term entry
        search_frame = ttk.Frame(params_frame)
        search_frame.pack(fill=tk.X, pady=5)
        
        search_label = ttk.Label(search_frame, text="Search Term:")
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_term_entry = ttk.Entry(search_frame, width=30)
        self.search_term_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Output options section
        output_frame = ttk.LabelFrame(main_frame, text="Output Options", padding=15)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Output type selection with radio buttons
        output_type_frame = ttk.Frame(output_frame)
        output_type_frame.pack(fill=tk.X, pady=5)
        
        self.output_option = tk.StringVar(value="new_file")
        
        new_file_radio = ttk.Radiobutton(output_type_frame, text="Create new Excel file", 
                                       variable=self.output_option, value="new_file")
        new_file_radio.pack(anchor=tk.W, pady=2)
        
        same_file_radio = ttk.Radiobutton(output_type_frame, text="Add to existing Excel file as new sheet", 
                                        variable=self.output_option, value="same_file")
        same_file_radio.pack(anchor=tk.W, pady=2)
        
        # Processing section
        actions_frame = ttk.Frame(main_frame, padding=15)
        actions_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_button = ttk.Button(actions_frame, text="Process Data", command=self.process_data)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        # Progress indicator
        self.progress_var = tk.IntVar()
        self.progress = ttk.Progressbar(actions_frame, variable=self.progress_var, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        self.quit_button = ttk.Button(actions_frame, text="Quit", command=self.master.quit)
        self.quit_button.pack(side=tk.RIGHT, padx=5)
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(status_frame, height=8, width=50, wrap=tk.WORD)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        self.status_text.insert(tk.END, "Welcome to Data Filter Application.\nPlease select required files and enter parameters to begin.\n")
        self.status_text.config(state=tk.DISABLED)

    def add_to_status(self, message):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"\n{message}")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.master.update_idletasks()

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_file_path = file_path
            filename = os.path.basename(file_path)
            self.excel_status.set(filename)
            self.add_to_status(f"Excel file selected: {filename}")
            
            # Try to load columns from the selected sheet
            self.load_columns()
        else:
            messagebox.showwarning("Warning", "No Excel file selected.")

    def load_columns(self):
        if not self.excel_file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
            
        sheet_name = self.sheet_name_entry.get().strip()
        if not sheet_name:
            messagebox.showwarning("Warning", "Please enter a sheet name.")
            return
            
        try:
            self.add_to_status(f"Loading columns from sheet: {sheet_name}")
            
            # Load the Excel file and get the column names
            df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
            self.available_columns = df.columns.tolist()
            
            # Update the dropdown with the column names
            self.filter_column_dropdown['values'] = self.available_columns
            
            if self.available_columns:
                # Select a default column (DiseaseName if available, otherwise first column)
                if 'DiseaseName' in self.available_columns:
                    self.filter_column.set('DiseaseName')
                else:
                    self.filter_column.set(self.available_columns[0])
                
                self.add_to_status(f"Loaded {len(self.available_columns)} columns from sheet: {sheet_name}")
            else:
                self.add_to_status("No columns found in the sheet.")
                
        except Exception as e:
            self.add_to_status(f"Error loading columns: {str(e)}")
            
            # Check if the sheet exists
            try:
                available_sheets = pd.ExcelFile(self.excel_file_path).sheet_names
                self.add_to_status(f"Available sheets: {', '.join(available_sheets)}")
            except:
                pass

    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_file_path = file_path
            filename = os.path.basename(file_path)
            self.pdf_status.set(filename)
            self.add_to_status(f"PDF file selected: {filename}")
        else:
            messagebox.showwarning("Warning", "No PDF file selected.")

    def process_data(self):
        # Get parameters
        filter_column = self.filter_column.get()
        search_term = self.search_term_entry.get()
        sheet_name = self.sheet_name_entry.get().strip() or "Sheet1"
        output_option = self.output_option.get()
        
        # Validate inputs
        if not self.excel_file_path or not self.pdf_file_path:
            messagebox.showwarning("Warning", "Please select both Excel and PDF files.")
            return
            
        if not filter_column:
            messagebox.showwarning("Warning", "Please select a column to filter by.")
            return
            
        if not search_term:
            messagebox.showwarning("Warning", "Please enter a search term.")
            return

        # Start progress bar
        self.progress.start()
        self.process_button.config(state="disabled")
        self.add_to_status(f"Processing data where {filter_column} contains '{search_term}' in sheet: {sheet_name}")
        
        try:
            self.add_to_status("Reading Excel file...")
            try:
                df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
            except ValueError as sheet_error:
                if "Worksheet named" in str(sheet_error) and "not found" in str(sheet_error):
                    available_sheets = pd.ExcelFile(self.excel_file_path).sheet_names
                    sheet_list = ", ".join(available_sheets)
                    self.add_to_status(f"Sheet '{sheet_name}' not found. Available sheets: {sheet_list}")
                    raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {sheet_list}")
                else:
                    raise sheet_error
                    
            if filter_column not in df.columns:
                raise ValueError(f"Column '{filter_column}' not found in the sheet.")

            # Filter the data based on the selected column and search term
            filtered_df = df[df[filter_column].astype(str).str.contains(search_term, case=False, na=False)]
            
            if filtered_df.empty:
                self.add_to_status(f"No data found where {filter_column} contains '{search_term}'.")
                self.progress.stop()
                self.process_button.config(state="normal")
                return
                
            self.add_to_status(f"Found {len(filtered_df)} records where {filter_column} contains '{search_term}'")

            # Save the original data for later merging
            filtered_df = filtered_df.reset_index(drop=True)
            original_df = filtered_df.copy()

            # Read the PDF file
            self.add_to_status("Reading PDF file...")
            pdf_document = fitz.open(self.pdf_file_path)
            pdf_text = "\n".join([pdf_document.load_page(i).get_text() for i in range(pdf_document.page_count)])
            pdf_document.close()
            self.add_to_status("PDF data extracted successfully")

            # Prepare data for AI
            data_text = filtered_df.to_string(index=False)

            # Prepare the AI prompt
            self.add_to_status("Preparing AI analysis...")
            
            # Alternative approach with three categories (if needed)
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
            self.add_to_status("Sending request to Gemini AI...")
            model = genai.GenerativeModel("gemini-2.0-flash")
            response = model.generate_content(prompt)

            if response.text:
                self.add_to_status("Processing AI response...")
                json_text = self.extract_json(response.text)
                if not json_text:
                    raise ValueError("No valid JSON found in AI response.")

                response_json = json.loads(json_text)

                # Save JSON output
                self.add_to_status("Saving results to JSON file...")
                base_name = os.path.splitext(os.path.basename(self.excel_file_path))[0]
                default_json_name = f"{base_name}by{filter_column}_Analyzed.json"
                
                self.output_json_path = filedialog.asksaveasfilename(
                    defaultextension=".json", 
                    filetypes=[("JSON files", "*.json")],
                    title="Save JSON results",
                    initialfile=default_json_name
                )
                
                if self.output_json_path:
                    with open(self.output_json_path, "w", encoding="utf-8") as json_file:
                        json.dump(response_json, json_file, indent=4)
                    self.add_to_status(f"JSON saved to: {os.path.basename(self.output_json_path)}")
                    
                    # Convert the JSON to DataFrame
                    response_df = pd.DataFrame(response_json)
                    
                    # Handle Excel output based on selected option
                    if output_option == "new_file":
                        # Create a new Excel file
                        self.add_to_status("Creating new Excel file from results...")
                        
                        # Generate default filename
                        default_excel_name = f"{base_name}by{filter_column}_Analyzed.xlsx"
                        
                        self.output_excel_path = filedialog.asksaveasfilename(
                            defaultextension=".xlsx", 
                            filetypes=[("Excel files", "*.xlsx")],
                            title="Save as new Excel file",
                            initialfile=default_excel_name
                        )
                        
                        if self.output_excel_path:
                            analyzed_sheet_name = f"{sheet_name}_Analyzed"
                            response_df.to_excel(self.output_excel_path, sheet_name=analyzed_sheet_name, index=False)
                            self.add_to_status(f"Results saved to new file: {os.path.basename(self.output_excel_path)}")
                        else:
                            self.add_to_status("Excel file save cancelled.")
                            
                    else:  # Add as new sheet to existing file
                        self.add_to_status("Adding results as a new sheet to the existing Excel file...")
                        
                        # Ask where to save the updated Excel file
                        default_excel_name = f"{base_name}_Updated.xlsx"
                        
                        self.output_excel_path = filedialog.asksaveasfilename(
                            defaultextension=".xlsx", 
                            filetypes=[("Excel files", "*.xlsx")],
                            title="Save updated Excel file",
                            initialfile=default_excel_name
                        )
                        
                        if self.output_excel_path:
                            # If the selected path is different from the input file, create a copy first
                            if self.output_excel_path != self.excel_file_path:
                                self.add_to_status("Creating a copy of the original Excel file...")
                                # Read all sheets from the original file
                                with pd.ExcelFile(self.excel_file_path) as original_excel:
                                    with pd.ExcelWriter(self.output_excel_path) as writer:
                                        for sheet in original_excel.sheet_names:
                                            pd.read_excel(self.excel_file_path, sheet_name=sheet).to_excel(
                                                writer, sheet_name=sheet, index=False
                                            )
                            
                            # Now add the new sheet
                            analyzed_sheet_name = f"{sheet_name}_Analyzed"
                            
                            # Check if the sheet already exists
                            try:
                                with pd.ExcelFile(self.output_excel_path) as xls:
                                    existing_sheets = xls.sheet_names
                                    
                                counter = 1
                                original_name = analyzed_sheet_name
                                while analyzed_sheet_name in existing_sheets:
                                    analyzed_sheet_name = f"{original_name}_{counter}"
                                    counter += 1
                                    
                                # Write to the file
                                with pd.ExcelWriter(self.output_excel_path, mode='a', if_sheet_exists='replace') as writer:
                                    response_df.to_excel(writer, sheet_name=analyzed_sheet_name, index=False)
                                    
                                self.add_to_status(f"Results added as sheet '{analyzed_sheet_name}' to: {os.path.basename(self.output_excel_path)}")
                            except Exception as write_error:
                                self.add_to_status(f"Error writing to Excel file: {str(write_error)}")
                                self.add_to_status("Creating a new file instead...")
                                
                                # Fallback to creating a new file
                                response_df.to_excel(self.output_excel_path, sheet_name=analyzed_sheet_name, index=False)
                                self.add_to_status(f"Results saved to new file: {os.path.basename(self.output_excel_path)}")
                        else:
                            self.add_to_status("Excel file save cancelled.")
                    
                    self.add_to_status("Process completed successfully!")
                else:
                    self.add_to_status("JSON file save cancelled.")
            else:
                self.add_to_status("Error: AI response is empty.")
                
        except Exception as e:
            self.add_to_status(f"Error: {str(e)}")
        finally:
            # Stop progress bar
            self.progress.stop()
            self.process_button.config(state="normal")

    def extract_json(self, text):
        match = re.search(r'\[.*\]', text, re.DOTALL)
        return match.group(0) if match else None

if __name__ == "__main__":
    # Use ThemedTk for better looking UI
    root = ThemedTk(theme="equilux")  # You can use other themes like: 'breeze', 'equilux', 'arc', etc.
    app = DataFilterApp(root)
    root.mainloop()