# Data Filter GUI

This project provides a graphical user interface (GUI) for filtering data based on medical guidelines using Python. The application allows users to interact with an Excel file and a PDF document to extract relevant information and generate a structured JSON response.

## Project Structure

```
data-filter-gui
├── src
│   ├── DataScript.py      # Main logic for filtering data
│   ├── gui.py             # GUI implementation using Tkinter
│   └── __init__.py        # Marks the directory as a Python package
├── requirements.txt        # Project dependencies
└── README.md               # Project documentation
```

## Requirements

To run this project, you need to install the following dependencies:

- pandas
- PyMuPDF (fitz)
- google-generativeai
- Tkinter (usually included with Python installations)

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Usage

1. Clone the repository or download the project files.
2. Navigate to the project directory.
3. Install the required dependencies.
4. Run the GUI application:

```
python src/gui.py
```

5. Use the interface to select the Excel file and PDF document, input the necessary parameters, and view the filtered results.

## Contributing

Feel free to submit issues or pull requests if you have suggestions for improvements or new features.

## License

This project is open-source and available under the MIT License.