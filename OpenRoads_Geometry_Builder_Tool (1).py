#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenRoads Designer Geometry Builder XML Generator — GPI Theme

✓ Excel → XML (one Geometry per sheet; Geometry name = sheet name)
✓ Settings: Theme (dark/light), Input Units (feet/meters/rods/chains), Output Units (feet/meters), Bearing Format (DMS/Decimal)
✓ Drag & Drop (optional if tkinterdnd2 installed)
✓ Status bar live hints, splash screen, details log
✓ From Deed PDF tab
    - Accepts a deed PDF, extracts text (pdfplumber / PyMuPDF), OCR fallback via pytesseract if Tesseract is set
    - Parses Lines/Curves (bearings, distances, radius, arc length/chord length, chord bearing)
    - Editable preview grid
    - Save to Excel in your converter’s schema
    - Send to Converter (prefills Excel path on main tab)

This script is built to degrade gracefully if optional packages are missing:
- Core Excel features need: pandas, openpyxl
- Icons/images: pillow
- PDF text: pdfplumber or PyMuPDF (fitz)
- OCR fallback: pytesseract + Tesseract binary
- Drag & drop: tkinterdnd2
"""

import sys, math, re, datetime, shlex, io, traceback, argparse
from pathlib import Path

# ---------------------- THEME / BRAND ----------------------
GPI_GREEN  = "#0F3320"
GPI_GREY   = "#A2AAAD"
GPI_HL     = "#84BD00"

THEME_MODE = "dark"  # default

def set_theme(mode: str):
    global THEME_MODE, BG_DARK, PANEL_DARK, PANEL_BORDER, TEXT_LIGHT, TEXT_SOFT, CONSOLE_BG, CONSOLE_FG, STATUS_BG, STATUS_FG
    THEME_MODE = "dark" if str(mode).lower() != "light" else "light"
    if THEME_MODE == "dark":
        BG_DARK      = GPI_GREEN
        PANEL_DARK   = "#10271C"
        PANEL_BORDER = "#2A3F34"
        TEXT_LIGHT   = "#E8F0EB"
        TEXT_SOFT    = "#C6D2CB"
        CONSOLE_BG   = "#0D2218"
        CONSOLE_FG   = "#E3F2EA"
        STATUS_BG    = "#0E241A"
        STATUS_FG    = "#CFE3D7"
    else:
        BG_DARK      = "#F5F7F5"
        PANEL_DARK   = "#FFFFFF"
        PANEL_BORDER = GPI_GREY
        TEXT_LIGHT   = "#122016"
        TEXT_SOFT    = "#4A5B51"
        CONSOLE_BG   = "#F3F5F3"
        CONSOLE_FG   = "#1B2A20"
        STATUS_BG    = "#E6ECE6"
        STATUS_FG    = "#1F2E22"

set_theme(THEME_MODE)

# ---------------------- UNITS ----------------------
UNIT_TO_FEET = {"feet":1.0,"meters":3.280839895013123,"rods":16.5,"chains":66.0}
FEET_TO_UNIT = {"feet":1.0,"meters":0.3048}

def normalize_unit_token(tok: str, default_unit="feet") -> str:
    if not tok: return default_unit
    t = tok.strip().lower().rstrip(".")
    if t in ("ft","foot","feet","'", "ft.","feet.","foot."): return "feet"
    if t in ("m","meter","meters","metre","metres","m.","meters."): return "meters"
    if t in ("rod","rods","rd","rds","rd.","rds."): return "rods"
    if t in ("chain","chains","ch","chs","ch.","chs."): return "chains"
    return default_unit

def convert_value_units(val: float, from_unit: str, to_unit: str) -> float:
    if val is None: return None
    if from_unit == to_unit: return float(val)
    ft = float(val) * UNIT_TO_FEET[from_unit]
    return ft * FEET_TO_UNIT[to_unit]

# ---------------------- SAFE IMPORTS ----------------------
def _try_import(modname):
    try:
        return __import__(modname)
    except Exception:
        return None

pandas      = _try_import("pandas")
openpyxl    = _try_import("openpyxl")
pdfplumber  = _try_import("pdfplumber")
fitz        = _try_import("fitz")  # PyMuPDF
pytesseract = _try_import("pytesseract")

try:
    from PIL import Image, ImageTk
    HAVE_PIL = True
except Exception:
    Image = ImageTk = None
    HAVE_PIL = False

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

from xml.etree.ElementTree import Element, SubElement, ElementTree

# ---------------------- GLOBAL EXCEPTION HOOK ----------------------
def _excepthook(exc_type, exc, tb):
    """Show a friendly error dialog for unexpected exceptions."""
    # KeyboardInterrupt is raised on Windows when the console window hosting
    # the script is closed. Treat it as a normal shutdown instead of an error
    # so that the user is not shown a scary dialog on exit.
    if exc_type is KeyboardInterrupt:
        return

    err = "".join(traceback.format_exception(exc_type, exc, tb))
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Unexpected Error", "An unexpected error occurred:\n\n" + err)
        root.destroy()
    except Exception:
        pass
    print(err, file=sys.stderr)

sys.excepthook = _excepthook

# ---------------------- ANGLES / BEARINGS ----------------------
def dms_to_decimal(body_str: str) -> float:
    s = str(body_str).strip().upper()
    s = s.replace("DEG","°").replace("º","°").replace("′","'").replace("’","'").replace("`","'").replace("″",'"')
    m = re.findall(r"[0-9.]+", s)
    if not m: return 0.0
    deg = float(m[0]) if len(m)>0 else 0.0
    minutes = float(m[1]) if len(m)>1 else 0.0
    seconds = float(m[2]) if len(m)>2 else 0.0
    return deg + minutes/60.0 + seconds/3600.0

def parse_bearing_to_east_ccw_radians(bearing, fmt: str):
    if bearing is None: return 0.0
    s = str(bearing).strip().upper()
    s_clean = s.replace(" ","")
    m = re.match(r"^([NS])(.+)([EW])$", s_clean)
    if m:
        ns, body, ew = m.groups()
        if fmt == "dms":
            deg = dms_to_decimal(body)
        else:
            try: deg = float(re.sub(r"[^0-9.\-]","",body))
            except: deg = dms_to_decimal(body)
        if   ns=="N" and ew=="E": az_ncw = deg
        elif ns=="N" and ew=="W": az_ncw = (360.0-deg)%360.0
        elif ns=="S" and ew=="E": az_ncw = (180.0-deg)%360.0
        else:                      az_ncw = (180.0+deg)%360.0
        deg_east_ccw = (90.0 - az_ncw) % 360.0
        return math.radians(deg_east_ccw)
    try:
        deg_from_north_cw = float(s)
        deg_east_ccw = (90.0 - deg_from_north_cw) % 360.0
        return math.radians(deg_east_ccw)
    except Exception:
        return 0.0

def normalize_angle_rad(a: float) -> float:
    while a <= -math.pi: a += 2*math.pi
    while a >  math.pi:  a -= 2*math.pi
    return a

# ---------------------- ARC GEOMETRY ----------------------
def compute_arc_from_radius_length(R: float, L: float):
    if R<=0 or L<=0: return None
    delta = L/R
    chord = 2*R*math.sin(delta/2.0)
    tangent = R*math.tan(delta/2.0)
    external = R*(1.0/math.cos(delta/2.0)-1.0)
    middle_ordinate = R*(1.0-math.cos(delta/2.0))
    return dict(delta=delta, chord=chord, tangent=tangent, external=external,
                middle_ordinate=middle_ordinate, arc_length=L)

def compute_arc_from_radius_chord(R: float, C: float):
    if R<=0 or C<=0 or C>2*R: return None
    delta = 2.0*math.asin(C/(2.0*R))
    arc_length = R*delta
    tangent = R*math.tan(delta/2.0)
    external = R*(1.0/math.cos(delta/2.0)-1.0)
    middle_ordinate = R*(1.0-math.cos(delta/2.0))
    return dict(delta=delta, chord=C, tangent=tangent, external=external,
                middle_ordinate=middle_ordinate, arc_length=arc_length)

# ---------------------- EXCEL HELPERS ----------------------
def _require_excel_libs():
    if pandas is None or openpyxl is None:
        raise RuntimeError("Missing Excel libraries.\n\nPlease install:\n  pip install pandas openpyxl")

def load_all_sheet_names(xlsx_path: Path):
    _require_excel_libs()
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()

def load_sheet_df(xlsx_path: Path, sheet_name: str):
    _require_excel_libs()
    return pandas.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")

def pick_numeric(row, *candidates):
    norm = {re.sub(r"[\s_()]","",str(k)).strip().lower(): k for k in row.index}
    for label in candidates:
        key = re.sub(r"[\s_()]","",label).strip().lower()
        if key in norm:
            v = row.get(norm[key])
            if pandas.isna(v): continue
            try: return float(str(v).replace(",",""))
            except: continue
    return None

def pick_text(row, *candidates):
    norm = {re.sub(r"[\s_()]","",str(k)).strip().lower(): k for k in row.index}
    for label in candidates:
        key = re.sub(r"[\s_()]","",label).strip().lower()
        if key in norm:
            v = row.get(norm[key])
            if pandas.isna(v): continue
            return str(v)
    return None

# ---------------------- XML BUILDERS ----------------------
def indent_etree(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent_etree(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i

def append_geometry_definition(parent: Element,
                               courses_df,
                               name: str,
                               bearing_fmt: str,
                               in_units: str,
                               out_units: str):
    gdef = SubElement(parent, "GeometryBuilderDefinition", {"name": name})
    info = SubElement(gdef, "GeometryBuilderInfo", {
        "featureName": "", "begX": "0", "begY": "0", "endX": "0", "endY": "0",
        "endType": "BeginPoint", "closureType": "None", "isForceTangent": "False",
    })
    legs = SubElement(info, "GeometryBuilderLegs")

    if courses_df is None or courses_df.empty or "Type" not in courses_df.columns:
        return

    in_to_ft = UNIT_TO_FEET[in_units]
    ft_to_out = FEET_TO_UNIT[out_units]
    prev_direction = None

    for _, row in courses_df.iterrows():
        typ = str(row.get("Type", "")).strip().lower()

        if typ == "line":
            bearing = pick_text(row, "Bearing")
            direction_rad = parse_bearing_to_east_ccw_radians(bearing, bearing_fmt) if bearing else (prev_direction or 0.0)
            prev_direction = direction_rad

            dist_in = pick_numeric(row, "Distance","Distance ft","Distance (ft)","Distance (m)","Distance (feet)","Dist")
            dist_ft = (dist_in or 0.0) * in_to_ft
            dist_out = dist_ft * ft_to_out

            SubElement(legs, "GeometryBuilderLeg", {
                "direction": f"{direction_rad:.15f}",
                "distance": f"{dist_out:.6f}",
                "legType": "Line",
                "isClockwise": "False",
                "isChordDefinition": "False",
                "arcParam1": "Radius",
                "arcParam2": "Length",
                "arcRadius": "0",
                "arcLength": "0",
                "arcLengthChorded": "0",
                "arcDelta": "0",
                "arcTangent": "0",
                "arcExternal": "0",
                "arcMiddleOrdinate": "0",
                "arcChord": "0",
                "spiType": "Clothoid",
                "spiParam": "Length",
                "spiLength": "0",
                "spiTheta": "0",
                "spiConst": "0",
                "spiStartIsInfinite": "False",
                "spiStartRadius": "∞",
                "spiEndIsInfinite": "False",
                "spiEndRadius": "∞",
            })

        elif typ in ("arc","curve"):
            R_in = pick_numeric(row, "Radius","Radius ft","Radius (ft)","Radius (m)")
            L_in = pick_numeric(row, "Arc Length","Arc Length ft","Arc Length (ft)","Arc Length (m)")
            C_in = pick_numeric(row, "Chord Length","Chord Length ft","Chord Length (ft)","Chord Length (m)")

            R_ft = (R_in or 0.0) * UNIT_TO_FEET[in_units]
            L_ft = (L_in if L_in is not None else 0.0) * UNIT_TO_FEET[in_units] if L_in is not None else None
            C_ft = (C_in if C_in is not None else 0.0) * UNIT_TO_FEET[in_units] if C_in is not None else None

            has_arc = L_ft is not None and L_ft > 0
            has_chord = C_ft is not None and C_ft > 0
            if R_ft <= 0 or (not has_arc and not has_chord):
                continue

            chord_bearing = pick_text(row, "Chord Bearing","ChordBearing")
            chord_dir = parse_bearing_to_east_ccw_radians(chord_bearing, bearing_fmt) if chord_bearing else None
            direction_for_xml = chord_dir if chord_dir is not None else (prev_direction if prev_direction is not None else 0.0)

            if has_arc:
                params = compute_arc_from_radius_length(R_ft, L_ft)
                arc_param2 = "Length"
            else:
                params = compute_arc_from_radius_chord(R_ft, C_ft)
                arc_param2 = "Chord"
            if params is None:
                params = dict(delta=0.0, chord=(C_ft or 0.0), tangent=0.0, external=0.0,
                              middle_ordinate=0.0, arc_length=(L_ft or 0.0))

            is_clockwise = "False"
            if chord_dir is not None and prev_direction is not None:
                defl = normalize_angle_rad(chord_dir - prev_direction)
                is_clockwise = "True" if defl < 0 else "False"

            R_out = R_ft * FEET_TO_UNIT[out_units]
            arc_len_out = params['arc_length'] * FEET_TO_UNIT[out_units]
            chord_out = params['chord'] * FEET_TO_UNIT[out_units]
            tan_out = params['tangent'] * FEET_TO_UNIT[out_units]
            ext_out = params['external'] * FEET_TO_UNIT[out_units]
            mo_out = params['middle_ordinate'] * FEET_TO_UNIT[out_units]

            SubElement(legs, "GeometryBuilderLeg", {
                "direction": f"{direction_for_xml:.15f}",
                "distance": f"{arc_len_out:.6f}",
                "legType": "Arc",
                "isClockwise": is_clockwise,
                "isChordDefinition": "False",
                "arcParam1": "Radius",
                "arcParam2": arc_param2,
                "arcRadius": f"{R_out:.6f}",
                "arcLength": f"{arc_len_out:.15f}",
                "arcLengthChorded": f"{chord_out:.15f}",
                "arcDelta": f"{params['delta']:.15f}",
                "arcTangent": f"{tan_out:.15f}",
                "arcExternal": f"{ext_out:.15f}",
                "arcMiddleOrdinate": f"{mo_out:.15f}",
                "arcChord": f"{chord_out:.15f}",
                "spiType": "Clothoid",
                "spiParam": "Length",
                "spiLength": "0",
                "spiTheta": "0",
                "spiConst": "0",
                "spiStartIsInfinite": "False",
                "spiStartRadius": "∞",
                "spiEndIsInfinite": "False",
                "spiEndRadius": "∞",
            })
        else:
            continue

def convert_excel_to_xml_multi(input_xlsx: Path, output_xml: Path, logger=None,
                               bearing_fmt: str = "dms", input_units: str = "feet", output_units: str = "feet"):
    sheets = load_all_sheet_names(input_xlsx)
    if logger:
        logger(f"Workbook has {len(sheets)} sheet(s): {', '.join(sheets)}")
        logger(f"Settings → Bearing Format: {bearing_fmt.upper()}, Input Units: {input_units.title()}, Output Units: {output_units.title()}")

    root = Element("GeometryBuilderDefinitions")
    rows = lines = curves = 0

    for name in sheets:
        if logger: logger(f"Reading sheet: {name!r}")
        try:
            df = load_sheet_df(input_xlsx, name)
            if logger and not df.empty:
                logger(f"  Rows: {len(df)}  Columns: {list(df.columns)}")
        except Exception as e:
            if logger: logger(f"  Failed to read sheet '{name}': {e}. Including empty geometry.")
            df = pandas.DataFrame() if pandas else None

        if df is not None and not df.empty and "Type" in df.columns:
            tl = df["Type"].astype(str).str.strip().str.lower()
            lcnt = int((tl == "line").sum())
            ccnt = int(((tl == "curve") | (tl == "arc")).sum())
            rows += len(df); lines += lcnt; curves += ccnt
            if logger: logger(f"  Found: Lines={lcnt} Curves={ccnt}")
        else:
            if logger: logger("  No 'Type' column or no usable rows; geometry will be empty.")

        append_geometry_definition(root, df, name=name,
                                   bearing_fmt=bearing_fmt, in_units=input_units, out_units=output_units)
        if logger: logger(f"  Appended geometry: {name!r}")

    if logger: logger("Writing XML file…")
    tree = ElementTree(root)
    indent_etree(root)
    tree.write(output_xml, encoding="UTF-16", xml_declaration=True)
    if logger: logger(f"Wrote XML → {output_xml}")
    return {"sheets": len(sheets), "rows": rows, "lines": lines, "curves": curves}

# ---------------------- PDF EXTRACTION & PARSING ----------------------
def try_set_tesseract_cmd(custom_path: str = None):
    if pytesseract is None: return False
    if custom_path and Path(custom_path).exists():
        pytesseract.pytesseract.tesseract_cmd = str(custom_path); return True
    for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
        if Path(p).exists():
            pytesseract.pytesseract.tesseract_cmd = p; return True
    cmd = getattr(pytesseract.pytesseract, "tesseract_cmd", "")
    return Path(cmd).exists() if cmd else False

def extract_text_from_pdf(pdf_path: Path, logger=None) -> str:
    text = ""
    if pdfplumber is not None:
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text() or ""
                    if t.strip(): text += t + "\n"
            if text.strip():
                if logger: logger("Text extracted with pdfplumber."); return text
        except Exception as e:
            if logger: logger(f"pdfplumber failed: {e}")
    if fitz is not None:
        try:
            doc = fitz.open(str(pdf_path))
            for p in doc:
                t = p.get_text("text")
                if t.strip(): text += t + "\n"
            if text.strip():
                if logger: logger("Text extracted with PyMuPDF."); return text
        except Exception as e:
            if logger: logger(f"PyMuPDF text extraction failed: {e}")
    if fitz is not None and pytesseract is not None:
        try:
            doc = fitz.open(str(pdf_path))
            for p in doc:
                pix = p.get_pixmap(dpi=300)
                img_bytes = pix.tobytes("png")
                if HAVE_PIL:
                    image = Image.open(io.BytesIO(img_bytes))
                    t = pytesseract.image_to_string(image)
                else:
                    t = pytesseract.image_to_string(Image.open(io.BytesIO(img_bytes)))
                if t.strip(): text += t + "\n"
            if text.strip():
                if logger: logger("Text extracted via OCR (PyMuPDF + pytesseract)."); return text
        except Exception as e:
            if logger: logger(f"OCR fallback failed: {e}")
    else:
        if logger:
            if fitz is None: logger("OCR skipped: PyMuPDF not available.")
            if pytesseract is None: logger("OCR skipped: pytesseract not available.")
    return text

_CHAR_NORMALIZE_MAP = {
    "′": "'",
    "’": "'",
    "`": "'",
    "″": '"',
    "“": '"',
    "”": '"',
    "º": "°",
    "‐": "-",
    "–": "-",
    "—": "-",
}

_DEG_WORD_PATTERN = re.compile(r"(?i)\bDEG(?:REE|REES)?\b")
_MIN_WORD_PATTERN = re.compile(r"(?i)\bMIN(?:UTE|UTES)?\b")
_SEC_WORD_PATTERN = re.compile(r"(?i)\bSEC(?:OND|ONDS)?\b")
_NORTH_WORD_PATTERN = re.compile(r"(?i)\bNORTH(?:ERLY)?\b")
_SOUTH_WORD_PATTERN = re.compile(r"(?i)\bSOUTH(?:ERLY)?\b")
_EAST_WORD_PATTERN = re.compile(r"(?i)\bEAST(?:ERLY)?\b")
_WEST_WORD_PATTERN = re.compile(r"(?i)\bWEST(?:ERLY)?\b")
_THEN_WORD_PATTERN = re.compile(r"(?i)\bTHEN\b")
_UNIT_PUNCT_PATTERN = re.compile(r"(?i)(FEET|FT|M|METERS|CHAIN|CHAINS|CHS|ROD|RODS|RDS)[\.,](?=\s)")
_NUM_PUNCT_PATTERN = re.compile(r"(?<=\d)[\.,](?=\s)")
_SPACE_PATTERN = re.compile(r"[ \t]+")


def _clean_text_for_parsing_with_map(t: str):
    if not t:
        return "", []

    s = t
    mapping = list(range(len(t)))

    def _apply_char_map(char_map):
        nonlocal s, mapping
        chars = []
        new_map = []
        for ch, idx in zip(s, mapping):
            repl = char_map.get(ch, ch)
            for new_ch in repl:
                chars.append(new_ch)
                new_map.append(idx)
        s = "".join(chars)
        mapping = new_map

    def _apply_regex(pattern, repl_fn):
        nonlocal s, mapping
        new_chunks = []
        new_map = []
        last_end = 0
        for m in pattern.finditer(s):
            start, end = m.span()
            new_chunks.append(s[last_end:start])
            new_map.extend(mapping[last_end:start])
            repl_text = repl_fn(m)
            if repl_text:
                match_map = mapping[start:end]
                if not match_map:
                    fallback = mapping[start - 1] if start > 0 else (mapping[end] if end < len(mapping) else 0)
                    match_map = [fallback]
                new_chunks.append(repl_text)
                rep_map = []
                for i_char in range(len(repl_text)):
                    rep_map.append(match_map[min(i_char, len(match_map) - 1)])
                new_map.extend(rep_map)
            last_end = end
        new_chunks.append(s[last_end:])
        new_map.extend(mapping[last_end:])
        s = "".join(new_chunks)
        mapping = new_map

    _apply_char_map(_CHAR_NORMALIZE_MAP)
    _apply_regex(_DEG_WORD_PATTERN, lambda m: "°")
    _apply_regex(_MIN_WORD_PATTERN, lambda m: "'")
    _apply_regex(_SEC_WORD_PATTERN, lambda m: '"')
    _apply_regex(_NORTH_WORD_PATTERN, lambda m: "N")
    _apply_regex(_SOUTH_WORD_PATTERN, lambda m: "S")
    _apply_regex(_EAST_WORD_PATTERN, lambda m: "E")
    _apply_regex(_WEST_WORD_PATTERN, lambda m: "W")
    _apply_regex(_THEN_WORD_PATTERN, lambda m: "THENCE")
    _apply_regex(_UNIT_PUNCT_PATTERN, lambda m: m.group(1))
    _apply_regex(_NUM_PUNCT_PATTERN, lambda m: "")
    _apply_regex(_SPACE_PATTERN, lambda m: " ")
    return s, mapping


def clean_text_for_parsing(t: str) -> str:
    return _clean_text_for_parsing_with_map(t)[0]

def _to_float(s: str):
    if s is None: return None
    try: return float(str(s).replace(",",""))
    except: return None

# ---------- FIXED, SAFE, VERBOSE REGEXES ----------
_LINE_QD_PATTERN = re.compile(r"""
    \b
    (?:THENCE\s+)?(?:ALONG\s+)?(?:THE\s+)?        # optional prose
    ([NS])\s*                                      # N or S
    (
        [0-9]{1,3}
        (?:
            [°º]\s*\d{1,2}(?:['’]\s*\d{1,2}(?:"|”)? )?   # DMS
            |
            \d+(?:\.\d+)?                                 # or decimal
        )
    )
    \s*([EW])                                      # E or W
    (?:[^0-9]{0,80})?                               # brief prose before distance
    (?:\b(?:FOR\s+)?(?:A\s+)?(?:DIST(?:ANCE)?|LENGTH)\s+(?:OF\s+)?)?  # optional distance prose
    ([0-9,]+(?:\.\d+)?)                             # distance number
    \s*(FEET|FT|METERS?|M|CHAINS?|CHS?|RODS?|RDS?)  # units
    (?=[\s,\.])
    """, re.IGNORECASE | re.DOTALL | re.VERBOSE)

_LINE_AZ_PATTERN = re.compile(r"""
    \b
    (?:THENCE\s+)?(?:ALONG\s+)?(?:A\s+BEARING\s+OF\s+)?  # optional prose
    ([0-9]{1,3}(?:\.\d+)?)\s*(?:°|DEG(?:REES?)?)          # azimuth degrees
    .*?
    \b([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)  # distance + units
    (?=[\s,\.])
    """, re.IGNORECASE | re.DOTALL | re.VERBOSE)

_CURVE_PATTERN = re.compile(r"""
    \bCURVE\s+TO\s+THE\s+(RIGHT|LEFT)\b
    .*?
    \bRADIUS\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b
    (?: .*? \bARC\s+LENGTH\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b )?
    (?: .*? \bCHORD\s+(?:DIST(?:ANCE)?|LENGTH)\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b )?
    (?: .*? \bCHORD\s+BEARS?\s+(N|S)\s*([0-9]{1,3}(?:[°º]\s*\d{1,2}(?:['’]\s*\d{1,2}(?:"|”)? )?|\d+(?:\.\d+)?))\s*(E|W) )?
    """, re.IGNORECASE | re.DOTALL | re.VERBOSE)


def _parse_deed_text_entries(cleaned_text: str, assumed_unit: str):
    entries = []
    taken_spans = []

    for m in _CURVE_PATTERN.finditer(cleaned_text):
        rl, rad, rad_unit, arc_len, arc_unit, chord_len, chord_unit, cns, cbody, cew = m.groups()
        radius = _to_float(rad)
        r_unit = normalize_unit_token(rad_unit, default_unit=assumed_unit)
        arc = _to_float(arc_len) if arc_len else None
        arc_u = normalize_unit_token(arc_unit, default_unit=assumed_unit) if arc_unit else None
        chord = _to_float(chord_len) if chord_len else None
        chord_u = normalize_unit_token(chord_unit, default_unit=assumed_unit) if chord_unit else None
        chord_bearing = f"{cns} {cbody} {cew}".upper().replace("  "," ") if (cns and cbody and cew) else None
        start, end = m.span()
        entries.append((start, end, {
            "Type": "Curve", "Bearing": None, "Distance": None, "DistanceUnit": None,
            "Radius": radius, "RadiusUnit": r_unit, "Arc Length": arc, "ArcUnit": arc_u,
            "Chord Length": chord, "ChordUnit": chord_u, "Chord Bearing": chord_bearing
        }))
        taken_spans.append((start, end))

    def _is_within_taken(pos: int) -> bool:
        for s0, e0 in taken_spans:
            if s0 <= pos < e0:
                return True
        return False

    for m in _LINE_QD_PATTERN.finditer(cleaned_text):
        if _is_within_taken(m.start()):
            continue
        ns, body, ew, dist, unit = m.groups()
        bearing = f"{ns} {body} {ew}".upper().replace("  ", " ")
        unit_norm = normalize_unit_token(unit, default_unit=assumed_unit)
        entries.append((m.start(), m.end(), {
            "Type": "Line", "Bearing": bearing, "Distance": _to_float(dist), "DistanceUnit": unit_norm,
            "Radius": None, "RadiusUnit": None, "Arc Length": None, "ArcUnit": None,
            "Chord Length": None, "ChordUnit": None, "Chord Bearing": None
        }))

    for m in _LINE_AZ_PATTERN.finditer(cleaned_text):
        if _is_within_taken(m.start()):
            continue
        az_deg, dist, unit = m.groups()
        bearing = str(az_deg).strip()
        unit_norm = normalize_unit_token(unit, default_unit=assumed_unit)
        entries.append((m.start(), m.end(), {
            "Type": "Line", "Bearing": bearing, "Distance": _to_float(dist), "DistanceUnit": unit_norm,
            "Radius": None, "RadiusUnit": None, "Arc Length": None, "ArcUnit": None,
            "Chord Length": None, "ChordUnit": None, "Chord Bearing": None
        }))

    entries.sort(key=lambda tup: tup[0])
    ordered_entries = []
    seen_positions = set()
    for start, end, data in entries:
        if start in seen_positions:
            continue
        seen_positions.add(start)
        ordered_entries.append((start, end, data))
    return ordered_entries


def parse_deed_text_to_dataframe(text: str, assumed_unit: str = "feet"):
    if pandas is None:
        raise RuntimeError("pandas is required to build the parsed table. Please install:\n  pip install pandas")
    cleaned_text = clean_text_for_parsing(text)
    ordered_entries = _parse_deed_text_entries(cleaned_text, assumed_unit)
    ordered_rows = [data for _, _, data in ordered_entries]
    df = pandas.DataFrame(ordered_rows, columns=[
        "Type", "Bearing", "Distance", "DistanceUnit", "Radius", "RadiusUnit",
        "Arc Length", "ArcUnit", "Chord Length", "ChordUnit", "Chord Bearing"
    ])
    return df


def find_call_spans_in_text(text: str, assumed_unit: str = "feet"):
    cleaned_text, mapping = _clean_text_for_parsing_with_map(text)
    if not cleaned_text or not mapping:
        return []
    ordered_entries = _parse_deed_text_entries(cleaned_text, assumed_unit)
    spans = []
    mapping_len = len(mapping)
    for start, end, data in ordered_entries:
        if start >= mapping_len:
            continue
        orig_start = mapping[start]
        end_index = min(max(end - 1, start), mapping_len - 1)
        orig_end = mapping[end_index] + 1
        if orig_end <= orig_start:
            orig_end = orig_start + 1
        spans.append((orig_start, orig_end, data.get("Type")))
    return spans

def dataframe_to_excel_schema(df, input_units_setting: str):
    if pandas is None:
        raise RuntimeError("pandas is required to build the Excel sheet. Please install:\n  pip install pandas")
    cols = ["Type","Bearing","Distance (ft)","Radius (ft)","Arc Length (ft)","Chord Length (ft)","Chord Bearing"]
    out = pandas.DataFrame(columns=cols)

    def conv(val, unit):
        if val is None or unit is None: return None
        u = normalize_unit_token(unit, input_units_setting)
        return convert_value_units(val, u, input_units_setting)

    for _, r in df.iterrows():
        typ = str(r.get("Type","")).strip().title()
        row_out = {
            "Type": typ if typ in ("Line","Curve") else "",
            "Bearing": r.get("Bearing") if r.get("Bearing") not in (None,"") else "",
            "Distance (ft)": conv(r.get("Distance"), r.get("DistanceUnit")) if r.get("Distance") not in (None,"") else "",
            "Radius (ft)": conv(r.get("Radius"), r.get("RadiusUnit")) if r.get("Radius") not in (None,"") else "",
            "Arc Length (ft)": conv(r.get("Arc Length"), r.get("ArcUnit")) if r.get("Arc Length") not in (None,"") else "",
            "Chord Length (ft)": conv(r.get("Chord Length"), r.get("ChordUnit")) if r.get("Chord Length") not in (None,"") else "",
            "Chord Bearing": r.get("Chord Bearing") if r.get("Chord Bearing") not in (None,"") else "",
        }
        out.loc[len(out)] = row_out

    out = out.dropna(how="all").fillna(value="")
    return out

# ---------------------- UI, GRID, SETTINGS, etc. (unchanged from previous message) ----------------------
class Splash(tk.Toplevel):
    def __init__(self, master, duration_ms=1100, on_done=None):
        super().__init__(master)
        self.on_done = on_done
        self.duration_ms = max(300, int(duration_ms))
        self.overrideredirect(True)
        self.configure(bg=BG_DARK)
        w, h = 560, 260
        self.update_idletasks()
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        x = (sw - w)//2; y = int((sh - h)*0.35)
        self.geometry(f"{w}x{h}+{x}+{y}")
        card = tk.Frame(self, bg=PANEL_DARK, highlightthickness=1, highlightbackground=PANEL_BORDER)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        head = tk.Frame(card, bg=PANEL_DARK); head.pack(fill="x", padx=16, pady=(14,6))
        try:
            logo_path = Path(__file__).with_name("GPI-768x768.jpg")
            if HAVE_PIL and logo_path.exists():
                img = Image.open(logo_path).resize((44,44), Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(img)
                tk.Label(head, image=self._logo, bg=PANEL_DARK).pack(side="left", padx=(0,12))
        except Exception: pass
        tk.Label(head, text="OpenRoads Designer Geometry Builder XML Generator",
                 bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",14,"bold")).pack(side="left", anchor="w")
        tk.Label(card, text="Initializing…", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",10)).pack(anchor="w", padx=16, pady=(0,12))
        self.style = ttk.Style(self)
        try: self.style.theme_use("clam")
        except: pass
        trough = CONSOLE_BG if THEME_MODE=="dark" else "#E7ECE7"
        self.style.configure("GPI.Horizontal.TProgressbar", troughcolor=trough, background=GPI_HL,
                             bordercolor=PANEL_BORDER, lightcolor=GPI_HL, darkcolor=GPI_HL)
        pb_frame = tk.Frame(card, bg=PANEL_DARK); pb_frame.pack(fill="x", padx=16, pady=(6,10))
        self.pb = ttk.Progressbar(pb_frame, orient="horizontal", mode="determinate", length=480,
                                  style="GPI.Horizontal.TProgressbar")
        self.pb.pack(fill="x")
        self.percent_lbl = tk.Label(card, text="0%", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",10))
        self.percent_lbl.pack(anchor="e", padx=16)
        tk.Label(card, text="GPI • Excel → XML + Deed PDF parsing", bg=PANEL_DARK,
                 fg=TEXT_SOFT, font=("Segoe UI",9)).pack(anchor="w", padx=16, pady=(6,12))
        self.after(60, self._animate)
    def _animate(self):
        steps = 100; interval = max(10, self.duration_ms//steps)
        def tick(i=0):
            pct = min(i,100); self.pb["value"]=pct; self.percent_lbl.config(text=f"{pct}%")
            if pct>=100: self.after(80, self._finish)
            else: self.after(interval, lambda: tick(i+2))
        tick(0)
    def _finish(self):
        try: self.destroy()
        finally:
            if self.on_done: self.on_done()

class DetailsDialog(tk.Toplevel):
    def __init__(self, master, title="Details"):
        super().__init__(master)
        self.title(title); self.configure(bg=PANEL_DARK)
        self.geometry("1000x620"); self.minsize(900,560)
        self.transient(master); self.grab_set()
        tk.Label(self, text=title, bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=12, pady=(10,6))
        frame = tk.Frame(self, bg=PANEL_DARK); frame.pack(fill="both", expand=True, padx=12, pady=(0,8))
        self.text = tk.Text(frame, bg=CONSOLE_BG, fg=CONSOLE_FG, relief="flat", wrap="word",
                            font=("Consolas",10), highlightthickness=1, highlightbackground=PANEL_BORDER)
        scroll = tk.Scrollbar(frame, command=self.text.yview); self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        btns = tk.Frame(self, bg=PANEL_DARK); btns.pack(fill="x", padx=12, pady=(2,12))
        tk.Button(btns, text="Copy Log", command=self.copy_log,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left")
        tk.Button(btns, text="Save Log…", command=self.save_log,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left", padx=6)
        tk.Button(btns, text="Close", command=self.destroy,
                  bg=GPI_HL, fg=GPI_GREEN, relief="flat", padx=16, pady=6, cursor="hand2").pack(side="right")
        self.log(f"Started at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    def log(self, msg:str):
        self.text.insert("end", msg+"\n"); self.text.see("end"); self.update_idletasks()
    def copy_log(self):
        data = self.text.get("1.0","end-1c"); self.clipboard_clear(); self.clipboard_append(data)
    def save_log(self):
        default = f"OpenRoads_XML_Generator_Log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        p = filedialog.asksaveasfilename(parent=self, title="Save Log", initialfile=default,
                                         defaultextension=".txt", filetypes=[("Text files","*.txt")])
        if not p: return
        with open(p,"w",encoding="utf-8") as f: f.write(self.text.get("1.0","end-1c"))

class SettingsDialog(tk.Toplevel):
    def __init__(self, master, current_mode, current_units_in, current_units_out, current_bearing_fmt,
                 current_tesseract_path, on_apply):
        super().__init__(master)
        self.title("Settings"); self.configure(bg=PANEL_DARK); self.geometry("520x420")
        self.transient(master); self.grab_set(); self.on_apply = on_apply
        tk.Label(self, text="Appearance", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(12,6))
        self.mode_var = tk.StringVar(value=current_mode)
        box = tk.Frame(self, bg=PANEL_DARK); box.pack(anchor="w", padx=14, pady=(2,10))
        tk.Radiobutton(box, text="Dark mode", variable=self.mode_var, value="dark",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Radiobutton(box, text="Light mode", variable=self.mode_var, value="light",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Label(self, text="Units", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        g1 = tk.Frame(self, bg=PANEL_DARK); g1.pack(fill="x", padx=14, pady=(2,2))
        tk.Label(g1, text="Input Units:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.units_in_var = tk.StringVar(value=current_units_in); tk.OptionMenu(g1, self.units_in_var, "feet","meters","rods","chains").pack(side="left")
        g2 = tk.Frame(self, bg=PANEL_DARK); g2.pack(fill="x", padx=14, pady=(2,8))
        tk.Label(g2, text="Output Units:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.units_out_var = tk.StringVar(value=current_units_out); tk.OptionMenu(g2, self.units_out_var, "feet","meters").pack(side="left")
        tk.Label(self, text="Bearing Format", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        self.bearing_var = tk.StringVar(value=current_bearing_fmt)
        b = tk.Frame(self, bg=PANEL_DARK); b.pack(anchor="w", padx=14, pady=(2,10))
        tk.Radiobutton(b, text="Degrees–Minutes–Seconds (DMS)", variable=self.bearing_var, value="dms",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Radiobutton(b, text="Decimal Degrees", variable=self.bearing_var, value="decimal",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Label(self, text="OCR (Tesseract) — optional for scanned PDFs", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        o = tk.Frame(self, bg=PANEL_DARK); o.pack(fill="x", padx=14, pady=(2,4))
        tk.Label(o, text="tesseract.exe path:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.tesseract_var = tk.StringVar(value=current_tesseract_path or "")
        self.tess_entry = tk.Entry(o, textvariable=self.tesseract_var, bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                   relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER, width=40)
        self.tess_entry.pack(side="left", fill="x", expand=True)
        tk.Button(o, text="Browse…", command=self._browse_tess,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="left", padx=6)
        btns = tk.Frame(self, bg=PANEL_DARK); btns.pack(fill="x", padx=14, pady=(8,14))
        tk.Button(btns, text="Cancel", command=self.destroy,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="right")
        tk.Button(btns, text="Apply", command=self._apply,
                  bg=GPI_HL, fg=GPI_GREEN, relief="flat", padx=16, pady=6, cursor="hand2").pack(side="right", padx=6)
    def _browse_tess(self):
        p = filedialog.askopenfilename(title="Select tesseract.exe", filetypes=[("tesseract.exe","tesseract.exe"),("All files","*.*")])
        if not p: return
        self.tesseract_var.set(p)
    def _apply(self):
        self.on_apply(self.mode_var.get(), self.units_in_var.get(), self.units_out_var.get(),
                      self.bearing_var.get(), self.tesseract_var.get())
        self.destroy()

UNCERTAIN_WRAP_LEFT = "⟪"
UNCERTAIN_WRAP_RIGHT = "⟫"
REVIEW_COLUMNS = {"Bearing", "Distance (ft)", "Radius (ft)", "Arc Length (ft)", "Chord Length (ft)", "Chord Bearing"}
NUMERIC_REVIEW_COLUMNS = {"Distance (ft)", "Radius (ft)", "Arc Length (ft)", "Chord Length (ft)"}


def _wrap_uncertain_display(val: str) -> str:
    if val is None:
        return ""
    s = str(val)
    if not s:
        return s
    if s.startswith(UNCERTAIN_WRAP_LEFT) and s.endswith(UNCERTAIN_WRAP_RIGHT):
        return s
    return f"{UNCERTAIN_WRAP_LEFT}{s}{UNCERTAIN_WRAP_RIGHT}"


def _strip_uncertainty_markup(val: str) -> str:
    if val is None:
        return ""
    s = str(val)
    if s.startswith(UNCERTAIN_WRAP_LEFT) and s.endswith(UNCERTAIN_WRAP_RIGHT):
        return s[len(UNCERTAIN_WRAP_LEFT):-len(UNCERTAIN_WRAP_RIGHT)]
    return s


def _coerce_value_for_column(col_name: str, val: str):
    if val in (None, ""):
        return ""
    if col_name in NUMERIC_REVIEW_COLUMNS:
        cleaned = str(val).replace(",", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return cleaned
    return val


class EditableGrid(ttk.Treeview):
    def __init__(self, parent, columns, on_edit_commit, **kwargs):
        super().__init__(parent, columns=columns, show="headings", **kwargs)
        self.on_edit_commit = on_edit_commit
        for c in columns:
            self.heading(c, text=c); self.column(c, width=150, stretch=True, anchor="w")
        self._editor = None; self.bind("<Double-1>", self._begin_edit)
        # Row highlighting tags for extraction confidence/edits.
        self.tag_configure("uncertain", background="", foreground=TEXT_SOFT)
        self.tag_configure("edited", background="#C7F9CC", foreground="#123822")
    def _begin_edit(self, event):
        if self.identify("region", event.x, event.y) != "cell": return
        row_id = self.identify_row(event.y); col_id = self.identify_column(event.x)
        if not row_id or not col_id: return
        col_index = int(col_id.replace("#",""))-1
        bbox = self.bbox(row_id, col_id)
        if not bbox:
            return
        x,y,w,h = bbox; value = self.set(row_id, self["columns"][col_index])
        self._editor = tk.Entry(self, relief="flat"); self._editor.insert(0, value)
        self._editor.select_range(0,"end"); self._editor.focus_set()
        self._editor.place(x=x,y=y,width=w,height=h)
        self._editor.bind("<Return>", lambda e: self._commit(row_id, col_index))
        self._editor.bind("<Escape>", lambda e: self._cancel())
        self._editor.bind("<FocusOut>", lambda e: self._commit(row_id, col_index))
    def _commit(self, row_id, col_index):
        if not self._editor: return
        new_val = self._editor.get(); col_name = self["columns"][col_index]
        self.set(row_id, col_name, new_val); self._editor.destroy(); self._editor=None
        if self.on_edit_commit: self.on_edit_commit(row_id, col_name, new_val)
    def _cancel(self):
        if self._editor: self._editor.destroy(); self._editor=None

BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

class App(BaseTk):
    def __init__(self):
        global DND_AVAILABLE
        self._dnd_error = None
        self.dnd_enabled = DND_AVAILABLE
        try:
            super().__init__()
        except tk.TclError as exc:
            if DND_AVAILABLE:
                DND_AVAILABLE = False
                self.dnd_enabled = False
                self._dnd_error = str(exc)
                tk.Tk.__init__(self)
            else:
                raise
        self.title("OpenRoads Designer Geometry Builder XML Generator — GPI")
        self.geometry("1260x820"); self.minsize(1140,760); self.configure(bg=BG_DARK)
        try:
            icon_path = Path(__file__).with_name("GPI-768x768.jpg")
            if HAVE_PIL and icon_path.exists():
                img = Image.open(icon_path).resize((32,32), Image.LANCZOS)
                self._icon_img = ImageTk.PhotoImage(img); self.iconphoto(True, self._icon_img)
        except Exception: pass
        self.settings = {"theme":THEME_MODE,"units_in":"feet","units_out":"feet","bearing_fmt":"dms","tesseract_path":""}
        self.deed_df = pandas.DataFrame() if pandas else None
        self.deed_pdf_path = None; self.deed_last_saved_excel = None
        self.console = None
        self._log_history = ["Ready."]
        self.grid_row_states = {}
        self._edited_rows = set()
        self.withdraw(); self.after(10, lambda: Splash(self, duration_ms=1100, on_done=self._after_splash))
    def _after_splash(self): self.deiconify(); self._build_ui()
    def _build_ui(self):
        for w in list(self.winfo_children()):
            if isinstance(w, tk.Toplevel): continue
            w.destroy()
        header = tk.Frame(self, bg=BG_DARK, height=64); header.pack(side="top", fill="x"); header.pack_propagate(False)
        try:
            logo_img_path = Path(__file__).with_name("GPI-768x768.jpg")
            if HAVE_PIL and logo_img_path.exists():
                self._logo = ImageTk.PhotoImage(Image.open(logo_img_path).resize((40,40), Image.LANCZOS))
                tk.Label(header, image=self._logo, bg=BG_DARK).pack(side="left", padx=18, pady=12)
        except Exception:
            tk.Label(header, text="GPI", bg=BG_DARK, fg=TEXT_LIGHT, font=("Segoe UI",18,"bold")).pack(side="left", padx=18, pady=12)
        tk.Label(header, text="OpenRoads Designer Geometry Builder XML Generator",
                 bg=BG_DARK, fg=TEXT_LIGHT, font=("Segoe UI",18,"bold")).pack(side="left", padx=8)
        tk.Label(header, text="Excel → XML (one geometry per sheet)", bg=BG_DARK, fg=TEXT_SOFT, font=("Segoe UI",10)).pack(side="left", padx=16)
        self._settings_btn = tk.Button(header, text="⚙", command=self.open_settings, bg=BG_DARK, fg=TEXT_LIGHT, relief="flat",
                                       font=("Segoe UI Symbol",16), padx=10, pady=2, cursor="hand2",
                                       activebackground=BG_DARK, activeforeground=GPI_HL, bd=0)
        self._settings_btn.pack(side="right", padx=16)
        card = tk.Frame(self, bg=PANEL_DARK, highlightthickness=1, highlightbackground=PANEL_BORDER, bd=0)
        card.pack(padx=22, pady=(16,0), fill="both", expand=True)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("ModernNotebook.TNotebook",
                        background=PANEL_DARK,
                        borderwidth=0,
                        tabmargins=(14, 10, 14, 0))
        style.configure("ModernNotebook.Tab",
                        background=PANEL_DARK,
                        foreground=TEXT_SOFT,
                        font=("Segoe UI", 11, "bold"),
                        padding=(24, 12))
        style.map("ModernNotebook.Tab",
                   background=[("selected", BG_DARK), ("active", PANEL_BORDER)],
                   foreground=[("selected", TEXT_LIGHT), ("active", TEXT_LIGHT)])
        style.layout("ModernNotebook.Tab", [
            ("Notebook.tab", {"sticky": "nswe", "children": [
                ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                    ("Notebook.label", {"sticky": "nswe"})
                ]})
            ]})
        ])

        self.notebook = ttk.Notebook(card, style="ModernNotebook.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        tab_qc    = tk.Frame(self.notebook, bg=PANEL_DARK);   self._build_deed_qc_tab(tab_qc);   self.notebook.add(tab_qc, text="Review")
        tab_call  = tk.Frame(self.notebook, bg=PANEL_DARK);   self._build_call_tab(tab_call);    self.notebook.add(tab_call, text="Extract")
        tab_excel = tk.Frame(self.notebook, bg=PANEL_DARK); self._build_excel_tab(tab_excel); self.notebook.add(tab_excel, text="Export")
        console_frame = tk.Frame(card, bg=PANEL_DARK); console_frame.pack(fill="both", expand=True, padx=16, pady=(0,10))
        tk.Label(console_frame, text="Messages", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w")
        self.console = tk.Text(console_frame, height=10, bg=CONSOLE_BG, fg=CONSOLE_FG, relief="flat", font=("Consolas",10),
                               wrap="word", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.console.pack(fill="both", expand=True)
        self._render_log_history()
        self.status_bar = tk.Frame(self, bg=STATUS_BG, height=26); self.status_bar.pack(side="bottom", fill="x")
        self.status_lbl = tk.Label(self.status_bar, text="Tip: Use Settings (⚙) to set units, bearing format, and OCR path.",
                                   bg=STATUS_BG, fg=STATUS_FG, anchor="w", padx=10, font=("Segoe UI",9))
        self.status_lbl.pack(side="left", fill="x")
        self._bind_hint(self._settings_btn, "Open Settings (theme, units, bearing format, OCR path)")
        if self._dnd_error:
            msg = ("Drag & drop support could not be enabled because the tkdnd library is missing. "
                   "The app will continue without drag & drop.\n\nDetails: " + self._dnd_error)
            self.after(200, lambda: messagebox.showwarning("Drag & Drop Unavailable", msg, parent=self))
            self._log("Drag & Drop disabled — tkdnd library not found. Continuing without drag & drop support.")
            self._dnd_error = None
    def _build_excel_tab(self, parent):
        tk.Label(parent, text="Excel workbook (.xlsx)", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(16,6))
        in_row = tk.Frame(parent, bg=PANEL_DARK); in_row.pack(fill="x", padx=16, pady=4)
        self.in_var = getattr(self,"in_var", tk.StringVar())
        self.in_entry = tk.Entry(in_row, textvariable=self.in_var, font=("Segoe UI",10),
                                 bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                 relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.in_entry.pack(side="left", fill="x", expand=True)
        btn_in = self._secondary_button(in_row, "Browse…", self.browse_in); btn_in.pack(side="left", padx=(8,0))
        self._bind_hint(self.in_entry, "Drop an .xlsx here or click Browse…"); self._bind_hint(btn_in, "Pick an Excel workbook")
        if self.dnd_enabled:
            try:
                self.in_entry.drop_target_register(DND_FILES); self.in_entry.dnd_bind("<<Drop>>", self._on_drop_excel)
            except Exception: pass
        tk.Label(parent, text="Output XML", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(10,6))
        out_row = tk.Frame(parent, bg=PANEL_DARK); out_row.pack(fill="x", padx=16, pady=4)
        self.out_var = getattr(self,"out_var", tk.StringVar())
        self.out_entry = tk.Entry(out_row, textvariable=self.out_var, font=("Segoe UI",10),
                                  bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                  relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.out_entry.pack(side="left", fill="x", expand=True)
        btn_out = self._secondary_button(out_row, "Save as…", self.browse_out); btn_out.pack(side="left", padx=(8,0))
        self._bind_hint(self.out_entry, "Choose where to save the XML"); self._bind_hint(btn_out, "Pick an output path for the XML")
        info = ("Input requirements & notes:\n"
                "  • Each Excel sheet becomes a separate geometry named after the sheet.\n"
                "  • Columns per row:\n"
                "      – Type: Line or Curve (Arc accepted)\n"
                "      – Lines: Bearing, Distance\n"
                "      – Curves: Radius + (Arc Length OR Chord Length), Chord Bearing\n"
                "  • Bearing Format setting controls how bearings are read (DMS or Decimal Degrees)\n"
                "  • Input Units setting controls how distances in Excel are interpreted\n"
                "  • Output Units setting controls the distances written into the XML\n"
                "  • Output is pretty-printed UTF-16 XML ready for OpenRoads Geometry Builder\n"
                "\n"
                "Creator: Levi Masters  •  Theme: GPI  •  Tool: OpenRoads Designer Geometry Builder XML Generator\n")
        tk.Label(parent, text=info, justify="left", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",10)).pack(fill="x", padx=16, pady=(8,4))
        actions = tk.Frame(parent, bg=PANEL_DARK); actions.pack(fill="x", padx=16, pady=(6,12))
        btn_convert = self._cta_button(actions, "Convert"); btn_convert.pack(side="left"); btn_convert.configure(command=self.convert)
        self._bind_hint(btn_convert, "Convert the workbook to Geometry Builder XML")
    def _build_deed_qc_tab(self, parent):
        tk.Label(parent, text="Deed PDF", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(16,6))
        pdf_row = tk.Frame(parent, bg=PANEL_DARK); pdf_row.pack(fill="x", padx=16, pady=4)
        self.pdf_var = getattr(self,"pdf_var", tk.StringVar())
        self.pdf_entry = tk.Entry(pdf_row, textvariable=self.pdf_var, font=("Segoe UI",10),
                                  bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                  relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.pdf_entry.pack(side="left", fill="x", expand=True)
        btn_browse_pdf = self._secondary_button(pdf_row, "Browse…", self.browse_pdf); btn_browse_pdf.pack(side="left", padx=(8,0))
        self._bind_hint(self.pdf_entry, "Drop a deed PDF here or click Browse…"); self._bind_hint(btn_browse_pdf, "Pick a deed PDF")
        if self.dnd_enabled:
            try:
                self.pdf_entry.drop_target_register(DND_FILES); self.pdf_entry.dnd_bind("<<Drop>>", self._on_drop_pdf)
            except Exception: pass
        btns = tk.Frame(parent, bg=PANEL_DARK); btns.pack(fill="x", padx=16, pady=(10,6))
        self.btn_extract_text = self._cta_button(btns, "Extract Text"); self.btn_extract_text.pack(side="left"); self.btn_extract_text.configure(command=self.extract_deed_text)
        self._bind_hint(self.btn_extract_text, "Extract deed text into an editable preview")
        self.btn_highlight_calls = self._secondary_button(btns, "Highlight Calls", self.highlight_calls_preview); self.btn_highlight_calls.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_highlight_calls, "Analyze the deed text and highlight detected calls")
        btn_clear = self._secondary_button(btns, "Clear Text", self.clear_deed_text); btn_clear.pack(side="left", padx=(10,0))
        self._bind_hint(btn_clear, "Clear the editable deed text")
        tk.Label(parent, text="Editable Deed Text", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(8,4))
        text_frame = tk.Frame(parent, bg=PANEL_DARK); text_frame.pack(fill="both", expand=True, padx=16, pady=(0,12))
        self.deed_text = tk.Text(text_frame, wrap="word", font=("Consolas", 10),
                                 bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                 relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.deed_text.pack(side="left", fill="both", expand=True)
        text_scroll = tk.Scrollbar(text_frame, orient="vertical", command=self.deed_text.yview)
        text_scroll.pack(side="right", fill="y")
        self.deed_text.configure(yscrollcommand=text_scroll.set)
        self._configure_call_highlight_tags()
        info_msg = ("Extract the deed PDF to populate the text above.\n"
                    "Review and edit as needed before running call extraction from the next tab.")
        tk.Label(parent, text=info_msg, justify="left", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",10)).pack(anchor="w", padx=16, pady=(0,12))

    def _build_call_tab(self, parent):
        tk.Label(parent, text="Call Extraction", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(16,6))
        btns = tk.Frame(parent, bg=PANEL_DARK); btns.pack(fill="x", padx=16, pady=(4,6))
        self.btn_extract_calls = self._cta_button(btns, "Run Extraction"); self.btn_extract_calls.pack(side="left"); self.btn_extract_calls.configure(command=self.extract_calls_from_text)
        self._bind_hint(self.btn_extract_calls, "Parse the edited deed text into calls")
        self.btn_save_excel = self._secondary_button(btns, "Save Excel…", self.save_deed_excel); self.btn_save_excel.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_save_excel, "Save parsed courses to an Excel file (converter-ready)")
        self.btn_send_converter = self._secondary_button(btns, "Send to Converter", self.send_to_converter); self.btn_send_converter.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_send_converter, "Auto-fill the Excel path in the Excel → XML tab")
        self._deed_style = ttk.Style(parent)
        try: self._deed_style.theme_use("clam")
        except: pass
        trough = CONSOLE_BG if THEME_MODE=="dark" else "#E7ECE7"
        self._deed_style.configure("GPI.Small.Horizontal.TProgressbar", troughcolor=trough, background=GPI_HL,
                                   bordercolor=PANEL_BORDER, lightcolor=GPI_HL, darkcolor=GPI_HL)
        pb_frame = tk.Frame(parent, bg=PANEL_DARK); pb_frame.pack(fill="x", padx=16, pady=(0,6))
        self.pb_deed = ttk.Progressbar(pb_frame, orient="horizontal", mode="determinate", length=320, style="GPI.Small.Horizontal.TProgressbar")
        self.pb_deed.pack(side="left"); self.pb_deed["value"]=0; self.pb_deed["maximum"]=100
        tk.Label(parent, text="Preview / Edit Courses", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", padx=16, pady=(8,4))
        grid_toolbar = tk.Frame(parent, bg=PANEL_DARK); grid_toolbar.pack(fill="x", padx=16, pady=(0,4))
        self.btn_insert_row = self._secondary_button(grid_toolbar, "Insert Row", self.insert_deed_row); self.btn_insert_row.pack(side="left")
        self.btn_delete_row = self._secondary_button(grid_toolbar, "Delete Row", self.delete_deed_row); self.btn_delete_row.pack(side="left", padx=(8,0))
        self._bind_hint(self.btn_insert_row, "Insert a blank course below the selected row")
        self._bind_hint(self.btn_delete_row, "Delete the selected course row(s)")
        grid_frame = tk.Frame(parent, bg=PANEL_DARK); grid_frame.pack(fill="both", expand=True, padx=16, pady=(0,10))
        cols = ["Type","Bearing","Distance (ft)","Radius (ft)","Arc Length (ft)","Chord Length (ft)","Chord Bearing"]
        self.grid = EditableGrid(grid_frame, columns=cols, on_edit_commit=self._grid_edit_commit); self.grid.pack(fill="both", expand=True)
        vsb = tk.Scrollbar(grid_frame, orient="vertical", command=self.grid.yview); hsb = tk.Scrollbar(grid_frame, orient="horizontal", command=self.grid.xview)
        self.grid.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set); vsb.pack(side="right", fill="y"); hsb.pack(side="bottom", fill="x")
        self._configure_grid_tags()
        if pandas is None: self._log("pandas not installed — PDF parsing and Excel saving will be disabled until installed (pip install pandas).")
        if openpyxl is None: self._log("openpyxl not installed — saving Excel will be disabled (pip install openpyxl).")
        if pdfplumber is None and fitz is None: self._log("pdfplumber / PyMuPDF not installed — text extraction may fail (pip install pdfplumber pymupdf).")
        if pytesseract is None: self._log("pytesseract not installed — OCR fallback disabled (pip install pytesseract).")
        self._refresh_grid_from_df()
    # Deed actions & helpers
    def extract_deed_text(self):
        p = Path(self.pdf_var.get() or "")
        if not p or not p.exists() or p.suffix.lower() != ".pdf":
            messagebox.showerror("Missing or invalid PDF", "Please select a valid .pdf deed file."); return
        details = DetailsDialog(self, title="Deed Text Extraction Details"); logger = details.log
        if getattr(self, "pb_deed", None):
            self.pb_deed["value"] = 0
            self.pb_deed["value"] = 10
        self._log(f"Extracting text from: {p}")
        if self.settings.get("tesseract_path"):
            if try_set_tesseract_cmd(self.settings["tesseract_path"]): logger(f"Tesseract path set: {self.settings['tesseract_path']}")
            else: logger("Provided Tesseract path invalid; OCR may be skipped.")
        else:
            if try_set_tesseract_cmd(): logger("Tesseract auto-detected.")
            else: logger("Tesseract not found; OCR will be used only if installed/auto-detected.")
        try:
            txt = extract_text_from_pdf(p, logger=logger) or ""
            if getattr(self, "pb_deed", None): self.pb_deed["value"] = 55
            if not txt.strip():
                logger("No text could be extracted; editing area left blank.")
                messagebox.showwarning("No text","No text found in the PDF (no text layer and OCR unavailable).")
        except Exception as e:
            logger(f"Extraction error: {e}"); messagebox.showerror("Extraction error", str(e)); return
        if getattr(self, "deed_text", None):
            self.deed_text.delete("1.0", "end")
            if txt:
                self.deed_text.insert("1.0", txt)
        if getattr(self, "pb_deed", None): self.pb_deed["value"] = 100
        self.deed_pdf_path = p
        self._log("Deed text ready for QC. Review/edit before running call extraction.")
        if txt and txt.strip():
            self.highlight_calls_preview(quiet=True)

    def clear_deed_text(self):
        if getattr(self, "deed_text", None):
            self.deed_text.delete("1.0", "end")
            self.deed_text.tag_remove("call_line", "1.0", "end")
            self.deed_text.tag_remove("call_curve", "1.0", "end")
        if getattr(self, "pb_deed", None): self.pb_deed["value"] = 0
        if pandas is not None:
            self.deed_df = pandas.DataFrame()
        else:
            self.deed_df = None
        self._edited_rows = set()
        self._refresh_grid_from_df()

    def extract_calls_from_text(self):
        if pandas is None:
            messagebox.showerror("Missing dependency", "pandas is required for building the parsed table.\nInstall with: pip install pandas"); return
        txt_widget = getattr(self, "deed_text", None)
        deed_text = txt_widget.get("1.0", "end") if txt_widget else ""
        if not deed_text.strip():
            messagebox.showerror("Missing deed text", "Provide deed text (extract from PDF or paste it) before running call extraction.")
            return
        details = DetailsDialog(self, title="Call Extraction Details"); logger = details.log
        if getattr(self, "pb_deed", None):
            self.pb_deed["value"] = 0
            self.pb_deed["value"] = 20
        self._log("Running call extraction from deed text…")
        try:
            deed_units_default = self.settings["units_in"]
            df_parsed = parse_deed_text_to_dataframe(deed_text, assumed_unit=deed_units_default)
            logger(f"Parsed rows: {len(df_parsed)}")
            if getattr(self, "pb_deed", None): self.pb_deed["value"] = 65
        except Exception as e:
            logger(f"Parsing error: {e}"); messagebox.showerror("Parsing error", str(e)); return
        try:
            df_excel = dataframe_to_excel_schema(df_parsed, input_units_setting=self.settings["units_in"])
            self._edited_rows = set()
            self.deed_df = df_excel; self._refresh_grid_from_df()
            if getattr(self, "pb_deed", None): self.pb_deed["value"] = 100
            self._log(f"Parsed {len(df_excel)} course(s). Review/edit and Save Excel…")
            logger(f"Converted to Excel schema rows: {len(df_excel)}")
        except Exception as e:
            logger(f"Schema conversion error: {e}"); messagebox.showerror("Schema conversion error", str(e)); return

    def extract_deed(self):  # backwards compatibility alias
        self.extract_calls_from_text()

    def highlight_calls_preview(self, quiet=False):
        if not getattr(self, "deed_text", None):
            return
        self._configure_call_highlight_tags()
        self.deed_text.tag_remove("call_line", "1.0", "end")
        self.deed_text.tag_remove("call_curve", "1.0", "end")
        text_value = self.deed_text.get("1.0", "end-1c")
        if not text_value.strip():
            if not quiet:
                messagebox.showinfo("No deed text", "Provide deed text (extract from PDF or paste it) before highlighting calls.")
            return
        try:
            spans = find_call_spans_in_text(text_value, assumed_unit=self.settings.get("units_in", "feet"))
        except Exception as e:
            self._log(f"Call preview error: {e}")
            if not quiet:
                messagebox.showerror("Call preview error", str(e))
            return
        if not spans:
            self._log("Call preview: no call patterns detected.")
            if not quiet:
                messagebox.showinfo("No calls detected", "No call patterns were detected in the deed text.")
            return
        for start, end, typ in spans:
            if end <= start:
                continue
            tag = "call_curve" if (typ and str(typ).lower() == "curve") else "call_line"
            self.deed_text.tag_add(tag, f"1.0+{start}c", f"1.0+{end}c")
        self._log(f"Call preview: highlighted {len(spans)} potential call(s).")

    def _configure_call_highlight_tags(self):
        if not getattr(self, "deed_text", None):
            return
        if THEME_MODE == "dark":
            line_bg = "#1F4D32"
            curve_bg = "#2B4066"
        else:
            line_bg = "#C7F9CC"
            curve_bg = "#D6E4FF"
        try:
            self.deed_text.tag_configure("call_line", background=line_bg)
            self.deed_text.tag_configure("call_curve", background=curve_bg)
        except Exception:
            pass
    def _refresh_grid_from_df(self):
        if not getattr(self, "grid", None):
            return
        self.grid.delete(*self.grid.get_children())
        self.grid_row_states = {}
        if self.deed_df is None or self.deed_df.empty:
            self._configure_grid_tags()
            return
        for idx, r in self.deed_df.iterrows():
            vals = [r.get(c,"") for c in self.grid["columns"]]
            if pandas is not None: vals = [("" if (isinstance(v,float) and pandas.isna(v)) else v) for v in vals]
            display_vals = []
            populated_cols = set()
            for col_name, val in zip(self.grid["columns"], vals):
                text_val = "" if val in (None, "") else val
                if str(text_val).strip() != "":
                    populated_cols.add(col_name)
                display_vals.append(text_val)
            row_id = self.grid.insert("", "end", values=display_vals)
            edited = idx in self._edited_rows
            review_cols = {c for c in populated_cols if c in REVIEW_COLUMNS}
            state_auto = set() if edited else review_cols
            self.grid_row_states[row_id] = {"edited": edited, "auto": state_auto}
            self._apply_row_style(row_id)
        self._configure_grid_tags()
    def _grid_edit_commit(self, row_id, col_name, new_val):
        if self.deed_df is None or self.deed_df.empty: return
        try: idx = list(self.grid.get_children()).index(row_id)
        except Exception: return
        if col_name in self.deed_df.columns and 0 <= idx < len(self.deed_df):
            cleaned_val = _strip_uncertainty_markup(new_val)
            coerced_val = _coerce_value_for_column(col_name, cleaned_val)
            self.deed_df.iat[idx, self.deed_df.columns.get_loc(col_name)] = coerced_val
            self.grid.set(row_id, col_name, cleaned_val)
            self._edited_rows.add(idx)
        state = self.grid_row_states.get(row_id)
        if state is not None:
            state.setdefault("auto", set()).discard(col_name)
            state["edited"] = True
            self._apply_row_style(row_id)

    def insert_deed_row(self):
        if pandas is None:
            messagebox.showerror("Missing dependency", "Row editing requires pandas.\nInstall with: pip install pandas"); return
        cols = list(self.grid["columns"])
        if self.deed_df is None or not isinstance(self.deed_df, pandas.DataFrame):
            self.deed_df = pandas.DataFrame(columns=cols)
        selection = self.grid.selection()
        children = list(self.grid.get_children())
        insert_index = len(children)
        if selection:
            try:
                insert_index = list(children).index(selection[0]) + 1
            except ValueError:
                insert_index = len(children)
        blank_values = ["" for _ in cols]
        new_iid = self.grid.insert("", insert_index, values=blank_values)
        self.grid.selection_set(new_iid); self.grid.focus(new_iid); self.grid.see(new_iid)
        self.grid_row_states[new_iid] = {"edited": False, "auto": set()}
        self._apply_row_style(new_iid)
        new_row_df = pandas.DataFrame([{c: "" for c in cols}])
        if self.deed_df.empty:
            self.deed_df = new_row_df
        else:
            top = self.deed_df.iloc[:insert_index] if insert_index > 0 else self.deed_df.iloc[:0]
            bottom = self.deed_df.iloc[insert_index:]
            self.deed_df = pandas.concat([top, new_row_df, bottom], ignore_index=True)
        self._edited_rows = {i if i < insert_index else i+1 for i in self._edited_rows}

    def delete_deed_row(self):
        if pandas is None:
            messagebox.showerror("Missing dependency", "Row editing requires pandas.\nInstall with: pip install pandas"); return
        selection = self.grid.selection()
        if not selection:
            messagebox.showinfo("Select a row", "Please select one or more rows to delete."); return
        children = list(self.grid.get_children())
        indices = sorted((children.index(item) for item in selection if item in children), reverse=True)
        if not indices:
            return
        for idx in indices:
            iid = children[idx]
            self.grid.delete(iid)
            self.grid_row_states.pop(iid, None)
        if isinstance(self.deed_df, pandas.DataFrame) and not self.deed_df.empty:
            drop_idx = [i for i in indices if i < len(self.deed_df)]
            if drop_idx:
                self.deed_df = self.deed_df.drop(self.deed_df.index[drop_idx]).reset_index(drop=True)
                new_edited = set()
                shift_points = set(drop_idx)
                for old_idx in self._edited_rows:
                    if old_idx in shift_points:
                        continue
                    shift = sum(1 for d in drop_idx if d < old_idx)
                    new_idx = old_idx - shift
                    if new_idx >= 0:
                        new_edited.add(new_idx)
                self._edited_rows = new_edited
        remaining = self.grid.get_children()
        if remaining:
            next_idx = min(indices[-1], len(remaining)-1)
            next_iid = remaining[next_idx]
            self.grid.selection_set(next_iid); self.grid.focus(next_iid); self.grid.see(next_iid)
    def _configure_grid_tags(self):
        if not getattr(self, "grid", None):
            return
        try:
            if THEME_MODE == "dark":
                self.grid.tag_configure("uncertain", background="", foreground=TEXT_SOFT)
                self.grid.tag_configure("edited", background="#1F4D32", foreground="#DFF7E7")
            else:
                self.grid.tag_configure("uncertain", background="", foreground=TEXT_SOFT)
                self.grid.tag_configure("edited", background="#C7F9CC", foreground="#123822")
        except Exception:
            pass
        for row_id in list(self.grid_row_states.keys()):
            self._apply_row_style(row_id)
    def _apply_row_style(self, row_id):
        state = self.grid_row_states.get(row_id)
        if state is None or not getattr(self, "grid", None):
            return
        edited = bool(state.get("edited"))
        review_cols = set() if edited else set(state.get("auto") or set())
        tags = []
        if edited:
            tags.append("edited")
        elif review_cols:
            tags.append("uncertain")
        self.grid.item(row_id, tags=tags)
        for col_name in self.grid["columns"]:
            current_val = self.grid.set(row_id, col_name)
            base_val = _strip_uncertainty_markup(current_val)
            if base_val != current_val:
                self.grid.set(row_id, col_name, base_val)
    def save_deed_excel(self):
        if self.deed_df is None or self.deed_df.empty:
            messagebox.showerror("Nothing to save","No parsed courses to save. Extract first."); return
        if pandas is None or openpyxl is None:
            messagebox.showerror("Missing dependency","Saving Excel requires pandas and openpyxl.\nInstall:\n  pip install pandas openpyxl"); return
        default_name = "DeedCourses.xlsx"
        p = filedialog.asksaveasfilename(title="Save Excel as", defaultextension=".xlsx", initialfile=default_name,
                                         filetypes=[("Excel files","*.xlsx")])
        if not p: return
        try:
            cols = ["Type","Bearing","Distance (ft)","Radius (ft)","Arc Length (ft)","Chord Length (ft)","Chord Bearing"]
            df = self.deed_df.copy()
            for c in cols:
                if c not in df.columns: df[c] = ""
            df = df[cols]; df.to_excel(p, index=False, engine="openpyxl")
            self.deed_last_saved_excel = Path(p)
            self._log(f"Saved Excel → {p}"); messagebox.showinfo("Saved", f"Excel written:\n{p}")
        except Exception as e:
            self._log(f"Save error: {e}"); messagebox.showerror("Save error", str(e))
    def send_to_converter(self):
        if not self.deed_last_saved_excel or not self.deed_last_saved_excel.exists():
            messagebox.showwarning("Save first","Please save the Excel file first, then send it to the converter."); return
        self.in_var.set(str(self.deed_last_saved_excel)); self.out_var.set(str(self.deed_last_saved_excel.with_suffix(".xml")))
        try:
            tabs = self.notebook.tabs()
            if tabs:
                self.notebook.select(tabs[-1])
        except Exception:
            pass
        self._log("Excel path sent to converter tab.")
    def browse_in(self):
        p = filedialog.askopenfilename(title="Select Excel workbook", filetypes=[("Excel files","*.xlsx")])
        if not p: return
        self.in_var.set(p); self.out_var.set(str(Path(p).with_suffix(".xml"))); self._log(f"Selected workbook: {p}")
    def browse_out(self):
        p = filedialog.asksaveasfilename(title="Save XML as", defaultextension=".xml", filetypes=[("XML files","*.xml")])
        if not p: return
        self.out_var.set(p); self._log(f"Output path set: {p}")
    def browse_pdf(self):
        p = filedialog.askopenfilename(title="Select Deed PDF", filetypes=[("PDF files","*.pdf")])
        if not p: return
        self.pdf_var.set(p); self.deed_pdf_path = Path(p); self._log(f"Selected deed PDF: {p}")
    def _on_drop_excel(self, event):
        p = self._extract_path_from_dnd(event.data)
        if not p:
            return
        if Path(p).suffix.lower() != ".xlsx":
            messagebox.showerror("Unsupported file","Please drop an .xlsx workbook."); return
        self.in_var.set(p); self.out_var.set(str(Path(p).with_suffix(".xml"))); self._log(f"Dropped workbook: {p}")
    def _on_drop_pdf(self, event):
        p = self._extract_path_from_dnd(event.data)
        if not p:
            return
        if Path(p).suffix.lower() != ".pdf":
            messagebox.showerror("Unsupported file","Please drop a .pdf deed file."); return
        self.pdf_var.set(p); self.deed_pdf_path = Path(p); self._log(f"Dropped deed PDF: {p}")
    def _extract_path_from_dnd(self, data: str):
        paths = []
        if data:
            try:
                cleaned = data.replace("{",'"').replace("}",'"'); paths = shlex.split(cleaned)
            except Exception:
                paths = data.strip().split()
        return paths[0] if paths else None
    def open_settings(self):
        SettingsDialog(self, current_mode=self.settings["theme"], current_units_in=self.settings["units_in"],
                       current_units_out=self.settings["units_out"], current_bearing_fmt=self.settings["bearing_fmt"],
                       current_tesseract_path=self.settings["tesseract_path"], on_apply=self.apply_settings)
    def apply_settings(self, mode, units_in, units_out, bearing_fmt, tess_path):
        previous = dict(self.settings)
        self.settings.update({"theme":mode,"units_in":units_in,"units_out":units_out,"bearing_fmt":bearing_fmt,"tesseract_path":tess_path or ""})
        if tess_path:
            ok = try_set_tesseract_cmd(tess_path)
            self._log("Tesseract path set." if ok else "Tesseract path invalid or not found.")
        elif previous.get("tesseract_path"):
            # Path cleared
            self._log("Tesseract path cleared; OCR will rely on auto-detect.")
        if previous.get("theme") != mode:
            in_path = self.in_var.get() if hasattr(self,"in_var") else ""
            out_path = self.out_var.get() if hasattr(self,"out_var") else ""
            pdf_path = self.pdf_var.get() if hasattr(self,"pdf_var") else ""
            deed_text_value = self.deed_text.get("1.0", "end") if hasattr(self, "deed_text") else ""
            selected_tab = None
            if hasattr(self, "notebook"):
                try: selected_tab = self.notebook.index("current")
                except Exception: selected_tab = None
            set_theme(mode); self.configure(bg=BG_DARK); self._build_ui()
            self.in_var.set(in_path); self.out_var.set(out_path); self.pdf_var.set(pdf_path)
            if deed_text_value and hasattr(self, "deed_text"):
                self.deed_text.insert("1.0", deed_text_value)
                self.highlight_calls_preview(quiet=True)
            if selected_tab is not None:
                try: self.notebook.select(selected_tab)
                except Exception: pass
        self._log(f"Settings applied → Theme={mode}, Input Units={units_in}, Output Units={units_out}, Bearing Format={bearing_fmt}")
    def convert(self):
        try:
            in_path = Path(self.in_var.get() or ""); out_path = Path(self.out_var.get() or "")
            if not in_path: messagebox.showerror("Missing input","Please select an Excel workbook (.xlsx)."); return
            if not in_path.exists(): messagebox.showerror("File not found", f"Input file not found:\n{in_path}"); return
            if not out_path: messagebox.showerror("Missing output","Please choose an output XML path."); return
            details = DetailsDialog(self, title="Processing Details"); logger = details.log
            self._log("Converting… see Details for live log.")
            stats = convert_excel_to_xml_multi(in_path, out_path, logger=logger,
                                               bearing_fmt=self.settings["bearing_fmt"],
                                               input_units=self.settings["units_in"],
                                               output_units=self.settings["units_out"])
            logger(f"Summary: Sheets={stats['sheets']} Rows={stats['rows']} Lines={stats['lines']} Curves={stats['curves']}")
            self._log(f"Done → {out_path}")
            self._log(f"Sheets: {stats['sheets']}  Rows: {stats['rows']}  Lines: {stats['lines']}  Curves: {stats['curves']}")
            messagebox.showinfo("Success", f"XML written:\n{out_path}\n\nSheets: {stats['sheets']}\nRows: {stats['rows']}\nLines: {stats['lines']}\nCurves: {stats['curves']}")
        except Exception as e:
            self._log(f"Error: {e}"); messagebox.showerror("Error", str(e))
    # helpers
    def _secondary_button(self, parent, text, cmd):
        b = tk.Button(parent, text=text, command=cmd, font=("Segoe UI",10,"bold"),
                      bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                      fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                      activebackground="#2E4A3C" if THEME_MODE=="dark" else "#C7D0C7",
                      activeforeground=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                      relief="flat", cursor="hand2", padx=14, pady=6, bd=0, highlightthickness=0)
        self._add_hover(b, "#243B2F" if THEME_MODE=="dark" else "#DFE4DF", "#2E4A3C" if THEME_MODE=="dark" else "#C7D0C7")
        return b
    def _cta_button(self, parent, text):
        b = tk.Button(parent, text=text, font=("Segoe UI",11,"bold"), bg=GPI_HL, fg=GPI_GREEN,
                      activebackground="#74A800", activeforeground="white", relief="flat", cursor="hand2",
                      padx=18, pady=8, bd=0, highlightthickness=0)
        self._add_hover(b, GPI_HL, "#74A800"); return b
    def _add_hover(self, widget, base, hover):
        widget.bind("<Enter>", lambda e: widget.configure(bg=hover))
        widget.bind("<Leave>", lambda e: widget.configure(bg=base))
    def _bind_hint(self, widget, msg):
        widget.bind("<Enter>", lambda e: self._set_hint(msg))
        widget.bind("<Leave>", lambda e: self._set_hint(""))
    def _set_hint(self, msg):
        self.status_lbl.config(text=("Ready." if not msg else msg))
    def _render_log_history(self):
        if not getattr(self, "console", None):
            return
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        for line in self._log_history:
            self.console.insert("end", line + "\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def _log(self, msg):
        self._log_history.append(msg)
        console = getattr(self, "console", None)
        if not console:
            return
        try:
            # ``winfo_exists`` returns 1 while the underlying Tk widget is alive. During
            # a theme rebuild the old Text widget is destroyed, and trying to interact
            # with it raises a TclError which previously halted the UI rebuild and
            # prevented the other tabs from being recreated.
            if not int(console.winfo_exists()):
                raise tk.TclError
            console.configure(state="normal")
            console.insert("end", msg+"\n")
            console.see("end")
            console.configure(state="disabled")
        except tk.TclError:
            # If the widget vanished (e.g. while rebuilding for a theme change),
            # drop the stale reference; the next rebuild will create a fresh console
            # and ``_render_log_history`` will repopulate it from ``_log_history``.
            self.console = None

def _run_cli(argv):
    parser = argparse.ArgumentParser(
        description="Convert an OpenRoads Geometry Excel workbook to XML without launching the GUI."
    )
    parser.add_argument("input", type=Path, help="Path to the source Excel workbook (.xlsx)")
    parser.add_argument("output", type=Path, help="Path where the XML output should be written")
    parser.add_argument(
        "--bearing-format", "-b",
        choices=("dms", "decimal"),
        default="dms",
        help="Bearing format used in the workbook (default: dms)",
    )
    parser.add_argument(
        "--input-units", "-i",
        choices=tuple(UNIT_TO_FEET.keys()),
        default="feet",
        help="Distance units used in the workbook (default: feet)",
    )
    parser.add_argument(
        "--output-units", "-o",
        choices=tuple(FEET_TO_UNIT.keys()),
        default="feet",
        help="Distance units for the generated XML (default: feet)",
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress progress logging (a summary is still printed)",
    )

    args = parser.parse_args(argv)

    if not args.input.exists():
        raise FileNotFoundError(f"Input workbook not found: {args.input}")

    if args.output.parent and not args.output.parent.exists():
        args.output.parent.mkdir(parents=True, exist_ok=True)

    logger = None if args.quiet else print

    stats = convert_excel_to_xml_multi(
        args.input,
        args.output,
        logger=logger,
        bearing_fmt=args.bearing_format,
        input_units=args.input_units,
        output_units=args.output_units,
    )

    summary = (
        f"Wrote XML → {args.output}\n"
        f"Sheets: {stats['sheets']}  Rows: {stats['rows']}  "
        f"Lines: {stats['lines']}  Curves: {stats['curves']}"
    )
    print(summary)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else list(argv)

    if argv:
        try:
            _run_cli(argv)
        except Exception as exc:
            sys.stderr.write(f"Error: {exc}\n")
            sys.exit(1)
        return

    try:
        app = App()
        app.mainloop()
    except tk.TclError as e:
        sys.stderr.write("GUI could not be started. Details:\n" + str(e) + "\n")
        sys.stderr.write(
            "Run with CLI arguments instead:\n  python OpenRoads_Geometry_Builder_Tool.py <input.xlsx> <output.xml>\n"
        )
        sys.exit(1)
    except KeyboardInterrupt:
        # Allow Ctrl+C (or console window closure on Windows) to exit quietly.
        pass

if __name__ == "__main__":
    main()
