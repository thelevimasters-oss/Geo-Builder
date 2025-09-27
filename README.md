# Geo-Builder

OpenRoads Designer Geometry Builder XML Generator — GPI Theme. The tool converts Excel workbooks and deed PDFs into geometry XML suitable for Bentley OpenRoads. It includes conveniences such as drag-and-drop, a status console, and live hints.

## Installation

1. Clone this repository.
2. Install the Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. (Optional) Install the [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) binary and ensure it is on your system `PATH` if you want OCR fallback for scanned PDFs.

## Dependency overview

The application is designed to degrade gracefully if optional packages are not available. Install the dependencies that correspond to the features you plan to use:

| Dependency | Purpose |
|------------|---------|
| `pandas`, `openpyxl` | Required for all Excel import/export features.
| `pillow` | Enables themed icons and other image assets in the GUI.
| `pdfplumber` | Primary engine for extracting text from deed PDFs.
| `PyMuPDF` (`fitz`) | Alternate PDF text extraction engine for better coverage on some PDFs.
| `pytesseract` + Tesseract binary | OCR fallback for scanned PDFs lacking embedded text.
| `tkinterdnd2` | Adds drag-and-drop file support to the interface.

If a dependency is missing, the GUI will disable the corresponding feature while keeping the rest of the tool functional.

## Usage

Run the GUI directly with Python:

```bash
python OpenRoads_Geometry_Builder_Tool.py
```

Follow the on-screen prompts to load Excel geometry workbooks or deed PDFs, adjust settings, and export the generated XML files.
