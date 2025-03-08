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

class DataFilterApp:
    def __init__(self, master):
        self.master = master
        master.title("Data Filter Application")
        
        # Set window size and position
        window_width = 650
        window_height = 550
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

        # Configure API key
        self.gemini_api_key = 'AIzaSyB1UxpRE-UYeLVLlaIX8vlQ9WFHH-gBCnQ'
        genai.configure(api_key=self.gemini_api_key)

        # Create a style object
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
        
        header_label = ttk.Label(header_frame, text="Medical Data Filter Application", style="Header.TLabel")
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
        
        disease_frame = ttk.Frame(params_frame)
        disease_frame.pack(fill=tk.X, pady=5)
        
        disease_label = ttk.Label(disease_frame, text="Disease Substring:")
        disease_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.disease_name_entry = ttk.Entry(disease_frame, width=30)
        self.disease_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Sheet name entry
        sheet_frame = ttk.Frame(params_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        sheet_label = ttk.Label(sheet_frame, text="Excel Sheet Name:")
        sheet_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.sheet_name_entry = ttk.Entry(sheet_frame, width=30)
        self.sheet_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.sheet_name_entry.insert(0, "IPD DEC-2024")  

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
        else:
            messagebox.showwarning("Warning", "No Excel file selected.")

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
        disease_name_substring = self.disease_name_entry.get()
        if not self.excel_file_path or not self.pdf_file_path or not disease_name_substring:
            messagebox.showwarning("Warning", "Please select files and enter a disease name substring.")
            return

        # Start progress bar
        self.progress.start()
        self.process_button.config(state="disabled")
        self.add_to_status(f"Processing data for disease: {disease_name_substring}")
        
        # Call the existing DataScript logic here
        try:
            self.add_to_status("Reading Excel file...")
            df = pd.read_excel(self.excel_file_path, sheet_name='IPD DEC-2024')
            if 'DiseaseName' not in df.columns:
                raise ValueError("Column 'DiseaseName' not found in the sheet.")

            filtered_df = df[df['DiseaseName'].str.contains(disease_name_substring, case=False, na=False)]
            if filtered_df.empty:
                self.add_to_status(f"No data found for '{disease_name_substring}'.")
                self.progress.stop()
                self.process_button.config(state="normal")
                return
            self.add_to_status(f"Found {len(filtered_df)} records for '{disease_name_substring}'")

            self.add_to_status("Reading PDF file...")
            pdf_document = fitz.open(self.pdf_file_path)
            pdf_text = "\n".join([pdf_document.load_page(i).get_text() for i in range(pdf_document.page_count)])
            pdf_document.close()
            self.add_to_status("PDF data extracted successfully")

            data_text = filtered_df.to_string(index=False)

            self.add_to_status("Preparing AI analysis...")
            prompt = f"""
            Analyze the following filtered data related to "{disease_name_substring}" and provide insights based on the guidelines.

            {pdf_text}

            Filtered Data:
            {data_text}

            Provide the response in **JSON format** with the following structure:

            [
              {{
                "Patient ID": "01-606944",
                "Prescribed Medications": ["Tab POLYMALT", "Tab CALDREE"],
                "Meets Guidelines": false,
                "Notes on Compliance": "Cefspan (Cefixime) is not a first-line choice for UTI."
              }},
              {{
                "Patient ID": "01-519619",
                "Prescribed Medications": ["Tab ERYTHROCIN", "Tab DUPHASTON"],
                "Meets Guidelines": false,
                "Notes on Compliance": "Erythromycin is not a first-line treatment for UTI."
              }}
            ]
            """

            self.add_to_status("Sending request to Gemini AI...")
            model = genai.GenerativeModel("gemini-1.5-flash")
            response = model.generate_content(prompt)

            if response.text:
                self.add_to_status("Processing AI response...")
                json_text = self.extract_json(response.text)
                if not json_text:
                    raise ValueError("No valid JSON found in AI response.")

                response_json = json.loads(json_text)

                self.add_to_status("Saving results to JSON file...")
                self.output_json_path = filedialog.asksaveasfilename(
                    defaultextension=".json", 
                    filetypes=[("JSON files", "*.json")],
                    title="Save JSON results"
                )
                if self.output_json_path:
                    with open(self.output_json_path, "w", encoding="utf-8") as json_file:
                        json.dump(response_json, json_file, indent=4)

                    self.add_to_status("Creating Excel file from results...")
                    response_df = pd.DataFrame(response_json)
                    self.output_excel_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx", 
                        filetypes=[("Excel files", "*.xlsx")],
                        title="Save Excel results"
                    )
                    if self.output_excel_path:
                        response_df.to_excel(self.output_excel_path, index=False)
                        self.add_to_status("Process completed successfully!")
                    else:
                        self.add_to_status("Excel file save cancelled.")
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
    root = ThemedTk(theme="arc")  # You can use other themes like: 'breeze', 'equilux', 'arc', etc.
    app = DataFilterApp(root)
    root.mainloop()