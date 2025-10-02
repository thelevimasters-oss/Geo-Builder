# Geo-Builder

Geo-Builder is a desktop and command-line companion for PlotClonjurer AI users who need to turn survey spreadsheets and deeds into OpenRoads-ready geometry. The tool ships as a single Python application with supporting helpers that can be launched with a themed Tk GUI or invoked from the terminal for unattended conversions.

## Key capabilities

- **Spreadsheet → XML conversion** – Transform one or more worksheets into OpenRoads geometry XML with per-sheet geometries, bearing/unit conversions, and validation feedback in the log panel.
- **Configurable workspace** – Persist GUI settings such as theme, units, bearing format, Tesseract location, and preferred State Plane Coordinate System. Optional `pyproj` transforms export coordinates directly into the selected SPCS.
- **Interactive parcel preview** – When `matplotlib` is available the GUI renders a live parcel sketch, highlights problem calls, and lets you export a DXF parcel via `ezdxf`.
- **Deed PDF workbench** – Extract deed text with `pdfplumber`/`PyMuPDF`, fall back to OCR through `pytesseract` + `pdf2image`, highlight detected calls, and send curated results back into the converter or save them to Excel/CSV.
- **AI-assisted deed parsing** – Bundle regex heuristics with an optional spaCy NER model. Train models directly inside the app, monitor progress, generate HTML training reports, and reuse saved pipelines.
- **Graceful degradation** – Optional capabilities (icons, drag-and-drop, PDF parsing, OCR, AI, preview, DXF export) are skipped automatically when their dependencies are missing. Startup diagnostics attempt to install/download missing spaCy assets when possible.

## Installation

Geo-Builder targets Python 3.9 or newer and relies on Tkinter being available (included with most CPython builds).

1. Create and activate a virtual environment:

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

2. Install the core requirements for Excel conversion:

   ```bash
   pip install pandas openpyxl
   ```

3. Add optional extras based on the features you need:

   ```bash
   pip install pillow pdfplumber pymupdf pdf2image pytesseract tkinterdnd2 spacy spacy-lookups-data pyproj ezdxf matplotlib
   ```

   - Install Tesseract OCR from your platform's package manager when using OCR features. Set `GEO_BUILDER_TESSERACT`, `TESSERACT_PATH`, `TESSERACT_CMD`, or `TESSERACT_EXE` if the binary is not on `PATH`.
   - After installing spaCy, download the default English model once:

     ```bash
     python -m spacy download en_core_web_sm
     ```

## Usage

### Launch the GUI

```bash
python "OpenRoads_Geometry_Builder_Tool (1).py"
```

From the GUI you can configure defaults, drag-and-drop workbooks (when `tkinterdnd2` is installed), review live conversion logs, inspect parcel previews, export DXF parcels, and access the Deed PDF and Deed AI tabs.

### Command-line conversion

Run the script with arguments to skip the GUI and convert straight from Excel to XML:

```bash
python "OpenRoads_Geometry_Builder_Tool (1).py" INPUT.xlsx OUTPUT.xml \
    --bearing-format {dms|decimal} \
    --input-units {feet|meters|rods|chains} \
    --output-units {feet|meters} \
    [--quiet]
```

The command validates file paths, creates any missing output folders, applies requested unit conversions, and prints a summary counting sheets, rows, lines, and curves processed.

### Deed AI workflow

1. Use the **Deed PDF** tab to ingest a deed, run extraction/OCR, and review detected calls in the editable grid.
2. Optional: highlight uncertain calls, adjust values, and export the cleaned worksheet for reuse in the converter.
3. Configure an AI training library (paired PDFs and labeled Excel files) and start model training directly from the settings dialog.
4. Once trained, run the **Deed AI** tab to extract calls with the spaCy model, export detected spans to CSV, or feed them back into your spreadsheets.

## Optional dependency quick reference

| Feature | Dependency |
| --- | --- |
| Icons and theming assets | `Pillow` |
| Drag-and-drop file loading | `tkinterdnd2` |
| PDF text extraction | `pdfplumber` or `PyMuPDF` (`fitz`) |
| OCR fallback | `pytesseract`, Tesseract binary, and optionally `pdf2image` + Poppler |
| Parcel preview plot | `matplotlib` |
| DXF parcel export | `ezdxf` |
| Coordinate reprojection | `pyproj` |
| Deed AI (NER) | `spacy`, `spacy-lookups-data`, trained model in `deed_ner_model/` |

## Repository layout

```
.
├── OpenRoads_Geometry_Builder_Tool (1).py  # GUI + CLI entry point
├── deed_extractor.py                       # Deed text normalization and AI helpers
├── deed_ner_model/                         # Bundled spaCy model/metadata cache
├── tests/                                  # Pytest suite (doctests + extraction helpers)
└── README.md
```

## Testing

Run the automated checks with:

```bash
pytest
```

The tests exercise deed extraction helpers and doctest coverage for the supporting utilities.

## Troubleshooting

- On Windows, the application can auto-detect or prompt for the Tesseract binary. Use the settings dialog or environment variables if detection fails.
- If spaCy is installed without the English model, the startup dependency check will attempt to download it. Offline environments should install the model beforehand.
- Optional integrations are detected dynamically—features tied to missing libraries remain hidden or are disabled with explanatory hints.

## Contributing

Issues and pull requests are welcome! Before submitting, create a virtual environment, install the dependencies relevant to your change, exercise both the GUI and CLI flows, and run the test suite.

## License

This project has not declared a license. If you intend to use it commercially, contact the original authors or maintainers to clarify licensing terms.
