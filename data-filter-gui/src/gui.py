import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dotenv import load_dotenv  # type: ignore

# Import local modules
from ai_service import AIService
from file_utils import save_to_json, save_to_excel, read_excel_file, read_pdf_file

class DataFilterApp:
    def __init__(self, master):
        self.master = master
        master.title("AI Medical Data Analyzer Application")
        
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

        # Initialize the AI service
        self.ai_service = AIService(self)

        # Load environment variables and configure AI
        self._configure_ai()

        # Create UI components
        self.create_widgets()

    def _configure_ai(self):
        """Load API key and configure AI service."""
        # Load environment variables from .env file (if it exists)
        load_dotenv()

        # Get API key from environment variables
        self.gemini_api_key = os.environ.get("GEMINI_API_KEY")
        if not self.gemini_api_key:
            messagebox.showerror("API Key Error", "Gemini API key not found in environment variables. Please set GEMINI_API_KEY.")
        else:
            self.ai_service.configure_api(self.gemini_api_key)

    def create_widgets(self):
        """Create all UI components."""
        # Main container with padding
        main_frame = ttk.Frame(self.master, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_label = ttk.Label(main_frame, text="AI Medical Data Analyzer Application", font=("Helvetica", 16, "bold"))
        header_label.pack(pady=(0, 20))
        
        # Files section
        files_frame = ttk.LabelFrame(main_frame, text="Input Files", padding=15)
        files_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Excel file selection
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(excel_frame, text="Excel File:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.excel_status = tk.StringVar(value="No file selected")
        ttk.Label(excel_frame, textvariable=self.excel_status, foreground="gray").pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(excel_frame, text="Browse...", command=self.select_excel_file).pack(side=tk.RIGHT)
        
        # PDF file selection
        pdf_frame = ttk.Frame(files_frame)
        pdf_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(pdf_frame, text="PDF File:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.pdf_status = tk.StringVar(value="No file selected")
        ttk.Label(pdf_frame, textvariable=self.pdf_status, foreground="gray").pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(pdf_frame, text="Browse...", command=self.select_pdf_file).pack(side=tk.RIGHT)
        
        # Parameters section
        params_frame = ttk.LabelFrame(main_frame, text="Filter Parameters", padding=15)
        params_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Sheet name entry
        sheet_frame = ttk.Frame(params_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sheet_frame, text="Excel Sheet Name:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.sheet_name_entry = ttk.Entry(sheet_frame, width=30)
        self.sheet_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.sheet_name_entry.insert(0, "Sheet1")
        
        ttk.Button(sheet_frame, text="Load Columns", command=self.load_columns).pack(side=tk.RIGHT)
        
        # Filter column selection
        filter_col_frame = ttk.Frame(params_frame)
        filter_col_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(filter_col_frame, text="Filter Column:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.filter_column = tk.StringVar()
        self.filter_column_dropdown = ttk.Combobox(filter_col_frame, textvariable=self.filter_column, state="readonly", width=28)
        self.filter_column_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.filter_column_dropdown['values'] = ["<Load Excel file first>"]
        
        # Search term entry
        search_frame = ttk.Frame(params_frame)
        search_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_frame, text="Search Term:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_term_entry = ttk.Entry(search_frame, width=30)
        self.search_term_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Output options section
        output_frame = ttk.LabelFrame(main_frame, text="Output Options", padding=15)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Radiobutton(output_frame, text="Create new files", 
                       variable=self.output_option, value="new_file").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(output_frame, text="Add to existing files", 
                       variable=self.output_option, value="same_file").pack(anchor=tk.W, pady=2)
        
        # Processing section
        actions_frame = ttk.Frame(main_frame, padding=15)
        actions_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_button = ttk.Button(actions_frame, text="Process Data", command=self.process_data)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        # Progress indicator
        self.progress = ttk.Progressbar(actions_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        ttk.Button(actions_frame, text="Quit", command=self.master.quit).pack(side=tk.RIGHT, padx=5)
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(status_frame, height=8, width=50, wrap=tk.WORD)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        self.status_text.insert(tk.END, "Welcome to AI Medical Data Analyzer.\nPlease select required files and enter parameters to begin.\n")
        self.status_text.config(state=tk.DISABLED)

    def add_to_status(self, message):
        """Add a message to the status text box."""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"\n{message}")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.master.update_idletasks()

    def select_excel_file(self):
        """Handle Excel file selection."""
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

    def select_pdf_file(self):
        """Handle PDF file selection."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_file_path = file_path
            filename = os.path.basename(file_path)
            self.pdf_status.set(filename)
            self.add_to_status(f"PDF file selected: {filename}")
        else:
            messagebox.showwarning("Warning", "No PDF file selected.")

    def load_columns(self):
        """Load and display Excel columns."""
        if not self.excel_file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
            
        sheet_name = self.sheet_name_entry.get().strip()
        if not sheet_name:
            messagebox.showwarning("Warning", "Please enter a sheet name.")
            return
            
        self.add_to_status(f"Loading columns from sheet: {sheet_name}")
        
        try:
            # Load the Excel file and get the column names
            df, error = read_excel_file(self.excel_file_path, sheet_name)
            if error:
                self.add_to_status(f"Error loading columns: {error}")
                return
                
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

    def process_data(self):
        """Main processing function."""
        # Get parameters
        filter_column = self.filter_column.get()
        search_term = self.search_term_entry.get().lower().strip()
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
            # Read Excel file
            self.add_to_status("Reading Excel file...")
            df, error = read_excel_file(self.excel_file_path, sheet_name)
            if error:
                raise ValueError(error)
                
            if filter_column not in df.columns:
                raise ValueError(f"Column '{filter_column}' not found in the sheet.")

            # Filter the data based on the selected column and search term
            filtered_df = self.ai_service.ai_assisted_filter(df, filter_column, search_term)
            
            if filtered_df.empty:
                self.add_to_status(f"No data found where {filter_column} contains '{search_term}'.")
                return
                
            self.add_to_status(f"Found {len(filtered_df)} records where {filter_column} contains '{search_term}'")

            # Save the original data for later merging
            filtered_df = filtered_df.reset_index(drop=True)

            # Read the PDF file
            self.add_to_status("Reading PDF file...")
            pdf_text, error = read_pdf_file(self.pdf_file_path)
            if error:
                raise ValueError(f"Error reading PDF: {error}")
                
            self.add_to_status("PDF data extracted successfully")

            # Prepare data for AI
            data_text = filtered_df.to_string(index=False)

            # Generate AI response
            response_json = self.ai_service.analyze_data(
                search_term, filter_column, pdf_text, data_text
            )
            
            if not response_json:
                self.add_to_status("Failed to get analyzable response from AI.")
                return
                
            # Convert to DataFrame for Excel output
            import pandas as pd
            response_df = pd.DataFrame(response_json)
            
            # Generate base filename for outputs
            base_name = os.path.splitext(os.path.basename(self.excel_file_path))[0]
            
            # Save JSON output
            is_new_file = (output_option == "new_file")
            json_path, error = save_to_json(response_json, self.excel_file_path, filter_column, is_new_file)
            if error:
                self.add_to_status(f"JSON save issue: {error}")
                if "cancelled" in error.lower():
                    self.add_to_status("Operation cancelled.")
                    return
            else:
                self.output_json_path = json_path
                self.add_to_status(f"JSON data saved successfully to: {os.path.basename(json_path)}")
            
            # Save Excel output
            excel_path, error = save_to_excel(response_df, self.excel_file_path, sheet_name, filter_column, is_new_file)
            if error:
                self.add_to_status(f"Excel save issue: {error}")
                if "cancelled" in error.lower():
                    self.add_to_status("Operation cancelled.")
                    return
            else:
                self.output_excel_path = excel_path
                self.add_to_status(f"Excel data saved successfully to: {os.path.basename(excel_path)}")
            
            self.add_to_status("Process completed successfully!")
            
        except Exception as e:
            self.add_to_status(f"Error: {str(e)}")
        finally:
            # Stop progress bar
            self.progress.stop()
            self.process_button.config(state="normal")