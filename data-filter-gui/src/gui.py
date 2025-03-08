import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
import fitz
import google.generativeai as genai
import re

class DataFilterGUI:
    def __init__(self, master):
        self.master = master
        master.title("Data Filter GUI")

        self.excel_file_path = ""
        self.pdf_file_path = ""
        self.output_json_path = ""
        self.output_excel_path = ""

        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self.master, text="Data Filter Application")
        self.label.pack()

        self.select_excel_button = tk.Button(self.master, text="Select Excel File", command=self.select_excel_file)
        self.select_excel_button.pack()

        self.select_pdf_button = tk.Button(self.master, text="Select PDF File", command=self.select_pdf_file)
        self.select_pdf_button.pack()

        self.disease_name_label = tk.Label(self.master, text="Disease Name Substring:")
        self.disease_name_label.pack()

        self.disease_name_entry = tk.Entry(self.master)
        self.disease_name_entry.pack()

        self.run_button = tk.Button(self.master, text="Run Filter", command=self.run_filter)
        self.run_button.pack()

        self.output_label = tk.Label(self.master, text="")
        self.output_label.pack()

    def select_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.excel_file_path:
            self.output_label.config(text=f"Selected Excel File: {os.path.basename(self.excel_file_path)}")

    def select_pdf_file(self):
        self.pdf_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_file_path:
            self.output_label.config(text=f"Selected PDF File: {os.path.basename(self.pdf_file_path)}")

    def run_filter(self):
        disease_name_substring = self.disease_name_entry.get()
        if not self.excel_file_path or not self.pdf_file_path or not disease_name_substring:
            messagebox.showerror("Error", "Please select files and enter a disease name substring.")
            return

        # Call the filtering logic from DataScript.py
        try:
            self.filter_data(disease_name_substring)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def filter_data(self, disease_name_substring):
        # Import the DataScript logic here
        from DataScript import filter_data_logic  # Assuming the logic is encapsulated in a function

        # Call the function with the necessary parameters
        result = filter_data_logic(self.excel_file_path, self.pdf_file_path, disease_name_substring)

        if result:
            self.output_label.config(text="Filtering completed successfully.")
        else:
            self.output_label.config(text="No data found for the specified disease.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataFilterGUI(root)
    root.mainloop()