# Geo-Builder

Geo-Builder is a desktop companion for Bentley OpenRoads Designer users who need to build XML geometry files from spreadsheet data. The tool ships as a single Python application that can be launched with a themed Tk GUI or invoked from the command line for unattended conversions.

## Features

- **Excel → XML conversion** – Create a valid OpenRoads geometry XML file from one or more worksheets. Each sheet is treated as a separate geometry with the worksheet name used as the geometry name.
- **Configurable settings** – Control input/output units (feet, meters, rods, chains) and choose between DMS and decimal bearing formats.
- **Modern Tk interface** – Dark/light theming, drag-and-drop support (when `tkinterdnd2` is installed), live status hints, and an activity log console.
- **Deed PDF helper** – Parse bearings, distances, and curve data from a deed PDF, pre-analyze and highlight detected calls in the editable text, export the interpreted data to Excel, and feed it back into the converter. Text extraction uses `pdfplumber`/`PyMuPDF` with an optional Tesseract OCR fallback.
- **Graceful degradation** – Optional capabilities (icons, drag-and-drop, PDF parsing, OCR) are skipped automatically if their dependencies are unavailable.

## Requirements

- Python 3.9 or newer (Tkinter is required for the GUI).
- Mandatory libraries for Excel conversion: `pandas`, `openpyxl`.
- Optional libraries that enhance the experience:
  - `Pillow` – display bundled icons and images.
  - `pdfplumber` or `PyMuPDF` (`fitz`) – extract text from deed PDFs.
  - `pytesseract` + Tesseract binary – OCR fallback for scanned PDFs.
  - `tkinterdnd2` – drag-and-drop file input in the GUI.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\\Scripts\\activate
pip install pandas openpyxl  # install optional extras as needed
```

Install any optional integrations you need, for example:

```bash
pip install pillow pdfplumber pymupdf pytesseract tkinterdnd2
```

If you plan to use OCR, install Tesseract from your platform's package manager. The tool now auto-detects `tesseract` when it is on `PATH`, and you can still point to a custom binary from the settings dialog if needed.

## Usage

### Launch the GUI

```bash
python "OpenRoads_Geometry_Builder_Tool (1).py"
```

From the GUI you can configure defaults, browse for Excel workbooks, review conversion logs, and access the Deed PDF helper tab.

### Command-line conversion

Run the script with arguments to skip the GUI and convert straight from Excel to XML:

```bash
python "OpenRoads_Geometry_Builder_Tool (1).py" INPUT.xlsx OUTPUT.xml \
    --bearing-format {dms|decimal} \
    --input-units {feet|meters|rods|chains} \
    --output-units {feet|meters} \
    [--quiet]
```

The command validates that the source workbook exists and creates any missing folders for the target XML file. A summary reporting the number of sheets, rows, lines, and curves processed is printed when the conversion completes.

## Project structure

```
.
├── README.md
└── OpenRoads_Geometry_Builder_Tool (1).py
```

All functionality resides in a single script, making it easy to bundle or freeze with tools such as `pyinstaller` if you prefer a standalone executable.

## Contributing

Issues and pull requests are welcome! If you plan to contribute code, please:

1. Create a virtual environment and install the dependencies.
2. Run the script locally to verify your changes in both GUI and CLI modes.
3. Format and lint your code before opening a pull request.

## License

This project has not declared a license. If you intend to use it commercially, please contact the original authors or maintainers to clarify licensing terms.
