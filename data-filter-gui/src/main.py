"""
Main entry point for the AI Medical Data Analyzer Application.
"""

import tkinter as tk
from ttkthemes import ThemedTk  # type: ignore
from gui import DataFilterApp

def main():
    """Application entry point."""
    # Use ThemedTk for better looking UI
    root = ThemedTk(theme="equilux")  # You can use other themes like: 'breeze', 'equilux', 'arc', etc.
    app = DataFilterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()