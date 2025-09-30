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

import sys, math, re, datetime, shlex, io, traceback, argparse, json, configparser, random, csv, importlib, subprocess, shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

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

@dataclass
class ParsedCall:
    type: str
    bearing: Optional[str] = None
    distance: Optional[float] = None
    radius: Optional[float] = None
    arc_length: Optional[float] = None
    chord_length: Optional[float] = None
    chord_bearing: Optional[str] = None
    rotation: Optional[str] = None


@dataclass
class ParcelSegment:
    type: str
    start: Tuple[float, float]
    end: Tuple[float, float]
    start_rel: Tuple[float, float]
    end_rel: Tuple[float, float]
    bearing: Optional[str] = None
    distance: Optional[float] = None
    radius: Optional[float] = None
    rotation: Optional[str] = None
    delta: Optional[float] = None
    bulge: float = 0.0
    center: Optional[Tuple[float, float]] = None


@dataclass
class ProcessResult:
    points_rel: List[Tuple[float, float]]
    points_abs: List[Tuple[float, float]]
    segments: List[ParcelSegment]


def _scroll_units_from_event(event) -> int:
    """Translate a Tkinter mouse wheel event into scroll units."""
    delta = getattr(event, "delta", 0)
    if delta:
        if sys.platform == "darwin":
            return -delta
        return -1 if delta > 0 else 1
    num = getattr(event, "num", None)
    if num == 4:
        return -1
    if num == 5:
        return 1
    return 0


def _bind_mousewheel_scroll(target, container=None, orient="vertical"):
    """Enable mouse wheel scrolling for a widget and its descendants."""

    def _on_mousewheel(event):
        units = _scroll_units_from_event(event)
        if not units:
            return
        if orient == "horizontal":
            target.xview_scroll(units, "units")
        else:
            target.yview_scroll(units, "units")
        return "break"

    def _bind_recursive(widget):
        try:
            widget.bind("<MouseWheel>", _on_mousewheel, add="+")
            widget.bind("<Button-4>", _on_mousewheel, add="+")
            widget.bind("<Button-5>", _on_mousewheel, add="+")
        except Exception:
            return
        for child in getattr(widget, "winfo_children", lambda: [])():
            _bind_recursive(child)

    _bind_recursive(container or target)

SPCS_ZONES = {
    "Alabama East (EPSG:26929)": 26929,
    "Alabama West (EPSG:26930)": 26930,
    "Alaska Zone 1 (EPSG:26931)": 26931,
    "Alaska Zone 2 (EPSG:26932)": 26932,
    "Alaska Zone 3 (EPSG:26933)": 26933,
    "Alaska Zone 4 (EPSG:26934)": 26934,
    "Alaska Zone 5 (EPSG:26935)": 26935,
    "Alaska Zone 6 (EPSG:26936)": 26936,
    "Alaska Zone 7 (EPSG:26937)": 26937,
    "Alaska Zone 8 (EPSG:26938)": 26938,
    "Alaska Zone 9 (EPSG:26939)": 26939,
    "Alaska Zone 10 (EPSG:26940)": 26940,
    "Arizona East (EPSG:26948)": 26948,
    "Arizona Central (EPSG:26949)": 26949,
    "Arizona West (EPSG:26950)": 26950,
    "Arkansas North (EPSG:26951)": 26951,
    "Arkansas South (EPSG:26952)": 26952,
    "California Zone 1 (EPSG:26941)": 26941,
    "California Zone 2 (EPSG:26942)": 26942,
    "California Zone 3 (EPSG:26943)": 26943,
    "California Zone 4 (EPSG:26944)": 26944,
    "California Zone 5 (EPSG:26945)": 26945,
    "California Zone 6 (EPSG:26946)": 26946,
    "Colorado North (EPSG:26953)": 26953,
    "Colorado Central (EPSG:26954)": 26954,
    "Colorado South (EPSG:26955)": 26955,
    "Connecticut (EPSG:26956)": 26956,
    "Delaware (EPSG:26957)": 26957,
    "Florida East (EPSG:26958)": 26958,
    "Florida West (EPSG:26959)": 26959,
    "Florida North (EPSG:26960)": 26960,
    "Georgia East (EPSG:26966)": 26966,
    "Georgia West (EPSG:26967)": 26967,
    "Hawaii Zone 1 (EPSG:26961)": 26961,
    "Hawaii Zone 2 (EPSG:26962)": 26962,
    "Hawaii Zone 3 (EPSG:26963)": 26963,
    "Hawaii Zone 4 (EPSG:26964)": 26964,
    "Hawaii Zone 5 (EPSG:26965)": 26965,
    "Idaho East (EPSG:26968)": 26968,
    "Idaho Central (EPSG:26969)": 26969,
    "Idaho West (EPSG:26970)": 26970,
    "Illinois East (EPSG:26971)": 26971,
    "Illinois West (EPSG:26972)": 26972,
    "Indiana East (EPSG:26973)": 26973,
    "Indiana West (EPSG:26974)": 26974,
    "Iowa North (EPSG:26975)": 26975,
    "Iowa South (EPSG:26976)": 26976,
    "Kansas North (EPSG:26977)": 26977,
    "Kansas South (EPSG:26978)": 26978,
    "Kentucky North (EPSG:26979)": 26979,
    "Kentucky South (EPSG:26980)": 26980,
    "Louisiana North (EPSG:26981)": 26981,
    "Louisiana South (EPSG:26982)": 26982,
    "Louisiana Offshore (EPSG:32198)": 32198,
    "Maine East (EPSG:26983)": 26983,
    "Maine West (EPSG:26984)": 26984,
    "Maryland (EPSG:26985)": 26985,
    "Massachusetts Mainland (EPSG:26986)": 26986,
    "Massachusetts Island (EPSG:26987)": 26987,
    "Michigan North (EPSG:26988)": 26988,
    "Michigan Central (EPSG:26989)": 26989,
    "Michigan South (EPSG:26990)": 26990,
    "Minnesota North (EPSG:26991)": 26991,
    "Minnesota Central (EPSG:26992)": 26992,
    "Minnesota South (EPSG:26993)": 26993,
    "Mississippi East (EPSG:26994)": 26994,
    "Mississippi West (EPSG:26995)": 26995,
    "Missouri East (EPSG:26996)": 26996,
    "Missouri Central (EPSG:26997)": 26997,
    "Missouri West (EPSG:26998)": 26998,
    "Montana (EPSG:32100)": 32100,
    "Nebraska (EPSG:32104)": 32104,
    "Nevada East (EPSG:32108)": 32108,
    "Nevada Central (EPSG:32109)": 32109,
    "Nevada West (EPSG:32110)": 32110,
    "New Hampshire (EPSG:32111)": 32111,
    "New Jersey (EPSG:32112)": 32112,
    "New Mexico East (EPSG:32113)": 32113,
    "New Mexico Central (EPSG:32114)": 32114,
    "New Mexico West (EPSG:32115)": 32115,
    "New York East (EPSG:32116)": 32116,
    "New York Central (EPSG:32117)": 32117,
    "New York West (EPSG:32118)": 32118,
    "New York Long Island (EPSG:2263)": 2263,
    "North Carolina (EPSG:32119)": 32119,
    "North Dakota North (EPSG:32121)": 32121,
    "North Dakota South (EPSG:32122)": 32122,
    "Ohio North (EPSG:32123)": 32123,
    "Ohio South (EPSG:32124)": 32124,
    "Oklahoma North (EPSG:32125)": 32125,
    "Oklahoma South (EPSG:32126)": 32126,
    "Oregon North (EPSG:32127)": 32127,
    "Oregon South (EPSG:32128)": 32128,
    "Pennsylvania North (EPSG:32129)": 32129,
    "Pennsylvania South (EPSG:32130)": 32130,
    "Puerto Rico & Virgin Is. (EPSG:32161)": 32161,
    "Rhode Island (EPSG:32130)": 32130,
    "South Carolina (EPSG:32133)": 32133,
    "South Dakota North (EPSG:32134)": 32134,
    "South Dakota South (EPSG:32135)": 32135,
    "Tennessee (EPSG:32136)": 32136,
    "Texas North (EPSG:32137)": 32137,
    "Texas North Central (EPSG:32138)": 32138,
    "Texas Central (EPSG:32139)": 32139,
    "Texas South Central (EPSG:32140)": 32140,
    "Texas South (EPSG:32141)": 32141,
    "Utah North (EPSG:32142)": 32142,
    "Utah Central (EPSG:32143)": 32143,
    "Utah South (EPSG:32144)": 32144,
    "Vermont (EPSG:32145)": 32145,
    "Virginia North (EPSG:32146)": 32146,
    "Virginia South (EPSG:32147)": 32147,
    "Washington North (EPSG:32148)": 32148,
    "Washington South (EPSG:32149)": 32149,
    "West Virginia North (EPSG:32150)": 32150,
    "West Virginia South (EPSG:32151)": 32151,
    "Wisconsin North (EPSG:32152)": 32152,
    "Wisconsin Central (EPSG:32153)": 32153,
    "Wisconsin South (EPSG:32154)": 32154,
    "Wyoming East (EPSG:32155)": 32155,
    "Wyoming East Central (EPSG:32156)": 32156,
    "Wyoming West Central (EPSG:32157)": 32157,
    "Wyoming West (EPSG:32158)": 32158,
}


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
_DEPENDENCY_CACHE: Dict[str, Optional[Any]] = {}


def _try_import(modname: str, pip_name: Optional[str] = None) -> Optional[Any]:
    if modname in _DEPENDENCY_CACHE:
        return _DEPENDENCY_CACHE[modname]
    try:
        module = importlib.import_module(modname)
        _DEPENDENCY_CACHE[modname] = module
        return module
    except Exception:
        pip_pkg = pip_name or modname
        try:
            print(f"[dependency] Installing {pip_pkg} …")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_pkg])
            module = importlib.import_module(modname)
            _DEPENDENCY_CACHE[modname] = module
            return module
        except Exception as exc:
            print(f"[dependency] Failed to install {pip_pkg}: {exc}")
            _DEPENDENCY_CACHE[modname] = None
            return None

"""Install optional AI dependencies with:
    pip install spacy pdf2image pytesseract pillow pandas openpyxl
"""

pandas      = _try_import("pandas")
openpyxl    = _try_import("openpyxl")
pdfplumber  = _try_import("pdfplumber")
fitz        = _try_import("fitz", "pymupdf")  # PyMuPDF
pytesseract = _try_import("pytesseract")
ezdxf       = _try_import("ezdxf")
pyproj      = _try_import("pyproj")
spacy       = _try_import("spacy")
pdf2image   = _try_import("pdf2image")
if spacy is not None:
    try:
        from spacy.training import Example
    except Exception:
        Example = None
else:
    Example = None


def ensure_spacy_model(model_name: str = "en_core_web_sm") -> bool:
    if spacy is None:
        return False
    model_name = str(model_name)
    model_path = Path(model_name)
    if model_path.exists():
        try:
            spacy.load(model_name)
            return True
        except Exception as exc:
            print(f"[dependency] Failed to load spaCy model at {model_path}: {exc}")
            return False
    try:
        spacy.load(model_name)
        return True
    except OSError:
        try:
            print(f"[dependency] Downloading spaCy model {model_name} …")
            subprocess.check_call([sys.executable, "-m", "spacy", "download", model_name])
            spacy.load(model_name)
            return True
        except Exception as exc:
            print(f"[dependency] Failed to download spaCy model {model_name}: {exc}")
            return False
    except Exception:
        return False

if ezdxf is not None:
    try:
        from ezdxf.lldxf import const as dxf_const
    except Exception:
        dxf_const = None
else:
    dxf_const = None

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAVE_MPL = True
except Exception:
    Figure = None
    FigureCanvasTkAgg = None
    HAVE_MPL = False

HAVE_EZDXF = ezdxf is not None
HAVE_PYPROJ = pyproj is not None

_try_import("PIL", "Pillow")
try:
    from PIL import Image, ImageTk
    HAVE_PIL = True
except Exception:
    Image = ImageTk = None
    HAVE_PIL = False

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
if _try_import("tkinterdnd2") is not None:
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        DND_AVAILABLE = True
    except Exception:
        DND_AVAILABLE = False
else:
    DND_AVAILABLE = False

from xml.etree.ElementTree import Element, SubElement, ElementTree


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tip_window: Optional[tk.Toplevel] = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip_window or not self.text:
            return
        bbox = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else None
        x = (bbox[0] if bbox else 0) + self.widget.winfo_rootx() + 25
        y = (bbox[3] if bbox else 0) + self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify="left", background="#ffffe0",
                         relief="solid", borderwidth=1, font=("Segoe UI", 9))
        label.pack(ipadx=4, ipady=2)

    def hide(self, _event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

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
    simple_dirs = {
        "N": 0.0,
        "NE": 45.0,
        "E": 90.0,
        "SE": 135.0,
        "S": 180.0,
        "SW": 225.0,
        "W": 270.0,
        "NW": 315.0,
    }
    if s_clean in simple_dirs:
        deg_from_north_cw = simple_dirs[s_clean]
        deg_east_ccw = (90.0 - deg_from_north_cw) % 360.0
        return math.radians(deg_east_ccw)
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
    if pytesseract is None:
        return False

    candidate_paths: List[Path] = []
    if custom_path:
        candidate_paths.append(Path(custom_path))
    else:
        auto_path = shutil.which("tesseract") or shutil.which("tesseract.exe")
        if auto_path:
            candidate_paths.append(Path(auto_path))
        candidate_paths.extend([
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ])

    for candidate in candidate_paths:
        if candidate.exists():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return True

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


def ocr_pdf_with_pytesseract(pdf_path: Path, dpi: int = 300, tesseract_path: Optional[str] = None,
                             logger=None, progress_callback=None) -> str:
    """OCR a scanned deed PDF to text.

    The function first attempts to rasterize the PDF with pdf2image (Poppler).
    If Poppler binaries are unavailable it automatically falls back to using
    PyMuPDF to render page images, avoiding manual installation steps.
    """
    if pytesseract is None:
        raise RuntimeError("OCR requires pytesseract. Install with: pip install pytesseract pillow")
    if tesseract_path:
        try_set_tesseract_cmd(tesseract_path)

    last_pdf2image_error: Optional[Exception] = None
    total_pages: int = 0

    def iter_page_images():
        nonlocal last_pdf2image_error, total_pages
        if pdf2image is not None:
            try:
                images = pdf2image.convert_from_path(str(pdf_path), dpi=dpi)
                total_pages = len(images)
                if logger:
                    logger(f"Rendering {total_pages} page(s) via pdf2image for OCR.")
                for image in images:
                    yield image
                return
            except Exception as exc:
                last_pdf2image_error = exc
                if logger:
                    logger(f"pdf2image convert_from_path failed: {exc}")
                    if "poppler" in str(exc).lower():
                        logger("Poppler not detected; attempting PyMuPDF fallback.")
        if fitz is not None:
            try:
                doc = fitz.open(str(pdf_path))
            except Exception as exc:
                raise RuntimeError(f"PyMuPDF could not open {pdf_path.name}: {exc}") from exc
            total_pages = doc.page_count or 0
            if logger:
                logger(f"Rendering {total_pages} page(s) via PyMuPDF for OCR.")
            for page in doc:
                if not HAVE_PIL:
                    raise RuntimeError("Pillow is required for PyMuPDF OCR fallback but is not available.")
                pix = page.get_pixmap(dpi=dpi)
                image_bytes = io.BytesIO(pix.tobytes("png"))
                image = Image.open(image_bytes)
                image.load()
                yield image
            return
        err_msg = "OCR requires either pdf2image (with Poppler) or PyMuPDF to render pages."
        if last_pdf2image_error is not None:
            err_msg += f" pdf2image error: {last_pdf2image_error}"
        raise RuntimeError(err_msg)

    text_blocks: List[str] = []
    processed_pages = 0
    for page_number, image in enumerate(iter_page_images(), start=1):
        processed_pages = page_number
        try:
            if progress_callback:
                progress_callback(page_number - 1, total_pages or page_number, "ocr")
            page_text = pytesseract.image_to_string(image)
            if page_text:
                text_blocks.append(page_text)
                if logger:
                    logger(f"OCR page {page_number}: {len(page_text.strip())} characters")
        except Exception as exc:
            if logger:
                logger(f"OCR failed on page {page_number}: {exc}")
    if progress_callback and processed_pages:
        progress_callback(processed_pages, total_pages or processed_pages, "ocr")
    return "\n".join(block.strip() for block in text_blocks if block).strip()


def _parse_start_end(value) -> Optional[Tuple[int, int]]:
    if value in (None, ""):
        return None
    if pandas is not None:
        try:
            if pandas.isna(value):
                return None
        except Exception:
            pass
    if isinstance(value, (list, tuple)) and len(value) >= 2:
        try:
            return int(value[0]), int(value[1])
        except (TypeError, ValueError):
            return None
    if isinstance(value, str):
        nums = [int(n) for n in re.findall(r"-?\d+", value)]
        if len(nums) >= 2:
            return nums[0], nums[1]
    return None


def _find_span_in_text(text: str, snippet: str, used_ranges: List[Tuple[int, int]]) -> Optional[Tuple[int, int]]:
    if not snippet:
        return None
    lower_text = text.lower()
    lower_snippet = snippet.lower()
    search_start = 0
    while True:
        idx = lower_text.find(lower_snippet, search_start)
        if idx == -1:
            return None
        start, end = idx, idx + len(snippet)
        overlap = False
        for used_start, used_end in used_ranges:
            if start < used_end and end > used_start:
                overlap = True
                break
        if not overlap:
            return start, end
        search_start = idx + len(lower_snippet)


def load_deed_training_dataset(library_folder: Path, tesseract_path: Optional[str] = None, logger=None) -> List[Tuple[str, Dict[str, Any]]]:
    """Load OCR text and labeled deed call spans from a training library folder."""
    if pandas is None:
        raise RuntimeError("pandas is required to read labeled Excel files. Install with: pip install pandas openpyxl")
    dataset: List[Tuple[str, Dict[str, Any]]] = []
    pdf_paths = sorted(library_folder.glob("*.pdf"))
    if not pdf_paths:
        if logger:
            logger("No PDF files found in training folder.")
        return dataset
    for pdf_path in pdf_paths:
        excel_path = None
        stem = pdf_path.stem
        for candidate in [library_folder / f"{stem}_correct.xlsx", library_folder / f"{stem}.xlsx"]:
            if candidate.exists():
                excel_path = candidate
                break
        if excel_path is None:
            if logger:
                logger(f"Skipping {pdf_path.name}: matching Excel file not found.")
            continue
        if logger:
            logger(f"OCR → {pdf_path.name}")
        text = ocr_pdf_with_pytesseract(pdf_path, tesseract_path=tesseract_path, logger=logger)
        if not text:
            if logger:
                logger(f"No OCR text produced for {pdf_path.name}; skipping.")
            continue
        try:
            df = pandas.read_excel(excel_path, engine="openpyxl")
        except Exception as exc:
            if logger:
                logger(f"Failed to read {excel_path.name}: {exc}")
            continue
        entities: List[Tuple[int, int, str]] = []
        used_ranges: List[Tuple[int, int]] = []
        for _, row in df.iterrows():
            call_text = str(row.get("Deed_Call_Text", "")).strip()
            if not call_text:
                continue
            start_end = _parse_start_end(row.get("Start_End"))
            if start_end is None:
                start_end = _find_span_in_text(text, call_text, used_ranges)
            if start_end is None:
                if logger:
                    snippet = call_text[:60] + ("…" if len(call_text) > 60 else "")
                    logger(f"Could not align call '{snippet}' in OCR text; skipping annotation.")
                continue
            start, end = start_end
            if start < 0 or end > len(text) or start >= end:
                continue
            used_ranges.append((start, end))
            entities.append((start, end, "DEED_CALL"))
        if not entities:
            if logger:
                logger(f"No valid annotations produced for {pdf_path.name}; skipping document.")
            continue
        dataset.append((text, {"entities": entities}))
        if logger:
            logger(f"Prepared {len(entities)} labeled call(s) from {pdf_path.name}.")
    return dataset


def train_deed_spacy_model(dataset: List[Tuple[str, Dict[str, Any]]],
                           output_dir: Optional[Path] = None,
                           base_model: str = "en_core_web_sm",
                           logger=None):
    """Train or fine-tune a spaCy NER model for deed calls."""
    if spacy is None or Example is None:
        raise RuntimeError("spaCy with training support is required. Install with: pip install spacy")
    if not dataset:
        raise ValueError("No training data provided for model training.")
    try:
        if base_model:
            if ensure_spacy_model(base_model):
                nlp = spacy.load(base_model)
            else:
                raise OSError
        else:
            raise OSError
    except Exception:
        nlp = spacy.blank("en")
        if logger:
            logger("Falling back to blank English spaCy model.")
    if "ner" not in nlp.pipe_names:
        ner = nlp.add_pipe("ner")
    else:
        ner = nlp.get_pipe("ner")
    if "DEED_CALL" not in ner.labels:
        ner.add_label("DEED_CALL")

    examples: List[Example] = []
    for text, annotations in dataset:
        doc = nlp.make_doc(text)
        spans = []
        for start, end, label in annotations.get("entities", []):
            span = doc.char_span(start, end, label=label)
            if span is not None:
                spans.append(span)
        if not spans:
            continue
        formatted = {"entities": [(span.start_char, span.end_char, span.label_) for span in spans]}
        examples.append(Example.from_dict(doc, formatted))
    if not examples:
        raise ValueError("Training data did not produce any valid spaCy examples.")

    other_pipes = [p for p in nlp.pipe_names if p != "ner"]
    with nlp.disable_pipes(*other_pipes):
        optimizer = nlp.initialize(lambda: examples)
        epochs = 10
        for epoch in range(epochs):
            random.shuffle(examples)
            losses = {}
            for batch in spacy.util.minibatch(examples, size=2):
                nlp.update(batch, drop=0.2, sgd=optimizer, losses=losses)
            if logger:
                loss_val = losses.get("ner", 0.0)
                logger(f"Epoch {epoch + 1}/{epochs} → NER loss={loss_val:.4f}")

    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
        nlp.to_disk(output_dir)
        if logger:
            logger(f"Model saved to {output_dir}")
    return nlp


def extract_deed_calls_with_model(nlp_model, text: str) -> List[Tuple[str, Tuple[int, int]]]:
    if nlp_model is None or not text:
        return []
    doc = nlp_model(text)
    calls: List[Tuple[str, Tuple[int, int]]] = []
    for ent in doc.ents:
        if ent.label_ == "DEED_CALL":
            calls.append((ent.text.strip(), (ent.start_char, ent.end_char)))
    return calls

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
_COMPOUND_CARDINAL_PATTERN = re.compile(r"(?i)\b(NORTH|SOUTH)(?:\s+|-)?(EAST|WEST)(?:ERLY)?\b")
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
    def _compound_cardinal_repl(m):
        text = m.group(0)
        suffix = "ERLY" if text.strip().upper().endswith("ERLY") else ""
        primary = m.group(1).upper()
        secondary = m.group(2).upper()
        return f"{primary}{secondary}{suffix}"
    _apply_regex(_COMPOUND_CARDINAL_PATTERN, _compound_cardinal_repl)
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

def _normalize_cardinal(token: Optional[str]) -> Optional[str]:
    if token is None:
        return None
    cleaned = str(token).strip().upper()
    if not cleaned:
        return None
    mapping = {
        "N": "N",
        "NORTH": "N",
        "S": "S",
        "SOUTH": "S",
        "E": "E",
        "EAST": "E",
        "W": "W",
        "WEST": "W",
    }
    return mapping.get(cleaned, cleaned[:1])


def _normalize_direction_word(token: Optional[str]) -> Optional[str]:
    if token is None:
        return None
    cleaned = re.sub(r"[^A-Z]", "", str(token).upper())
    if not cleaned:
        return None
    mapping = {
        "NORTHERLY": "N",
        "SOUTHERLY": "S",
        "EASTERLY": "E",
        "WESTERLY": "W",
        "NORTHEAST": "NE",
        "NORTHEASTERLY": "NE",
        "NORTHWEST": "NW",
        "NORTHWESTERLY": "NW",
        "SOUTHEAST": "SE",
        "SOUTHEASTERLY": "SE",
        "SOUTHWEST": "SW",
        "SOUTHWESTERLY": "SW",
    }
    return mapping.get(cleaned)

# ---------- FIXED, SAFE, VERBOSE REGEXES ----------
_LINE_QD_PATTERN = re.compile(r"""
    \b
    (?:THENCE\s+)?(?:ALONG\s+)?(?:THE\s+)?        # optional prose
    (N(?:ORTH)?|S(?:OUTH)?)\s*                      # N/S or NORTH/SOUTH
    (
        [0-9]{1,3}
        (?:
            [°º]\s*\d{1,2}(?:['’]\s*\d{1,2}(?:"|”)? )?   # DMS
            |
            \d+(?:\.\d+)?                                 # or decimal
        )
    )
    \s*(E(?:AST)?|W(?:EST)?)                       # E/W or EAST/WEST
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

_LINE_SIMPLE_DIR_PATTERN = re.compile(r"""
    \b
    (?:THENCE\s+)?(?:RUNNING\s+)?(?:CONTINUING\s+)?(?:ALONG\s+)?(?:THE\s+)?
    (?P<dir>(?:NORTH|SOUTH)(?:\s+|-)?(?:EAST|WEST)(?:ERLY)?|(?:EAST|WEST|NORTH|SOUTH)ERLY)
    (?P<prose>[^0-9]{0,80}?)
    (?:\b(?:FOR\s+)?(?:A\s+)?(?:DIST(?:ANCE)?|LENGTH)\s+(?:OF\s+)?)?
    (?P<dist>[0-9,]+(?:\.\d+)?)
    \s*(?P<unit>FEET|FT|METERS?|M|CHAINS?|CHS?|RODS?|RDS?)
    (?=[\s,\.])
    """, re.IGNORECASE | re.DOTALL | re.VERBOSE)

_CURVE_PATTERN = re.compile(r"""
    \bCURVE\s+TO\s+THE\s+(RIGHT|LEFT)\b
    .*?
    \bRADIUS\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b
    (?: .*? \bARC\s+LENGTH\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b )?
    (?: .*? \bCHORD\s+(?:DIST(?:ANCE)?|LENGTH)\s+(?:OF\s+)?([0-9,]+(?:\.\d+)?)\s*(FEET|FT|METERS?|M|CHAINS?|RODS?|RDS?)\b )?
    (?: .*? \bCHORD\s+BEARS?\s+(N(?:ORTH)?|S(?:OUTH)?)\s*([0-9]{1,3}(?:[°º]\s*\d{1,2}(?:['’]\s*\d{1,2}(?:"|”)? )?|\d+(?:\.\d+)?))\s*(E(?:AST)?|W(?:EST)?) )?
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
        if cns and cbody and cew:
            chord_ns = _normalize_cardinal(cns)
            chord_ew = _normalize_cardinal(cew)
            chord_bearing = (
                f"{chord_ns} {cbody} {chord_ew}".upper().replace("  ", " ")
                if (chord_ns and chord_ew)
                else None
            )
        else:
            chord_bearing = None
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
        ns_token = _normalize_cardinal(ns)
        ew_token = _normalize_cardinal(ew)
        bearing = (
            f"{ns_token} {body} {ew_token}".upper().replace("  ", " ")
            if (ns_token and ew_token)
            else None
        )
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

    for m in _LINE_SIMPLE_DIR_PATTERN.finditer(cleaned_text):
        if _is_within_taken(m.start()):
            continue
        direction_token = m.group("dir")
        prose = m.group("prose") or ""
        if re.search(r"\b(CURVE|RADIUS|ARC)\b", prose, re.IGNORECASE):
            continue
        bearing = _normalize_direction_word(direction_token)
        if not bearing:
            continue
        dist = m.group("dist")
        unit = m.group("unit")
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
        _bind_mousewheel_scroll(self.text)
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
                 current_tesseract_path, current_spcs, current_ai_dir, on_apply):
        super().__init__(master)
        self.title("Settings"); self.configure(bg=PANEL_DARK); self.geometry("560x520")
        self.transient(master); self.grab_set(); self.on_apply = on_apply
        self.ai_training_var = tk.StringVar(value=current_ai_dir or "")
        container = tk.Frame(self, bg=PANEL_DARK)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, bg=PANEL_DARK, highlightthickness=0, borderwidth=0)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        vscroll.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vscroll.set)
        content = tk.Frame(canvas, bg=PANEL_DARK)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        def _sync_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(window_id, width=event.width)
        content.bind("<Configure>", _sync_scroll)
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        _bind_mousewheel_scroll(canvas)
        _bind_mousewheel_scroll(canvas, content)
        tk.Label(content, text="Appearance", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(12,6))
        self.mode_var = tk.StringVar(value=current_mode)
        box = tk.Frame(content, bg=PANEL_DARK)
        box.pack(anchor="w", padx=14, pady=(2,10))
        tk.Radiobutton(box, text="Dark mode", variable=self.mode_var, value="dark",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Radiobutton(box, text="Light mode", variable=self.mode_var, value="light",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Label(content, text="Units", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        g1 = tk.Frame(content, bg=PANEL_DARK)
        g1.pack(fill="x", padx=14, pady=(2,2))
        tk.Label(g1, text="Input Units:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.units_in_var = tk.StringVar(value=current_units_in)
        tk.OptionMenu(g1, self.units_in_var, "feet", "meters", "rods", "chains").pack(side="left")
        g2 = tk.Frame(content, bg=PANEL_DARK)
        g2.pack(fill="x", padx=14, pady=(2,8))
        tk.Label(g2, text="Output Units:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.units_out_var = tk.StringVar(value=current_units_out)
        tk.OptionMenu(g2, self.units_out_var, "feet", "meters").pack(side="left")
        tk.Label(content, text="Coordinate System", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        cs_frame = tk.Frame(content, bg=PANEL_DARK)
        cs_frame.pack(fill="x", padx=14, pady=(2,8))
        tk.Label(cs_frame, text="SPCS Selection:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self._spcs_label = tk.Label(cs_frame, text=current_spcs or "None", bg=PANEL_DARK, fg=TEXT_LIGHT, anchor="w")
        self._spcs_label.pack(side="left", padx=(0,10))
        tk.Button(cs_frame, text="Choose…", command=self._choose_spcs,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="left")
        tk.Label(content, text="Bearing Format", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        self.bearing_var = tk.StringVar(value=current_bearing_fmt)
        b = tk.Frame(content, bg=PANEL_DARK)
        b.pack(anchor="w", padx=14, pady=(2,10))
        tk.Radiobutton(b, text="Degrees–Minutes–Seconds (DMS)", variable=self.bearing_var, value="dms",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Radiobutton(b, text="Decimal Degrees", variable=self.bearing_var, value="decimal",
                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT).pack(anchor="w")
        tk.Label(content, text="OCR (Tesseract) — optional for scanned PDFs", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(10,6))
        o = tk.Frame(content, bg=PANEL_DARK)
        o.pack(fill="x", padx=14, pady=(2,4))
        tk.Label(o, text="tesseract.exe path:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.tesseract_var = tk.StringVar(value=current_tesseract_path or "")
        self.tess_entry = tk.Entry(o, textvariable=self.tesseract_var, bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                                   relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER, width=40)
        self.tess_entry.pack(side="left", fill="x", expand=True)
        tk.Button(o, text="Browse…", command=self._browse_tess,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="left", padx=6)
        tk.Label(content, text="AI Tools", bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",12,"bold")).pack(anchor="w", padx=14, pady=(12,6))
        tk.Label(content, text="Manage the spaCy deed-call model and training library.", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",9), wraplength=460, justify="left").pack(anchor="w", padx=14, pady=(0,6))
        ai_path_row = tk.Frame(content, bg=PANEL_DARK)
        ai_path_row.pack(fill="x", padx=14, pady=(0,8))
        tk.Label(ai_path_row, text="Training folder:", bg=PANEL_DARK, fg=TEXT_SOFT, width=14, anchor="w").pack(side="left")
        self.ai_entry = tk.Entry(ai_path_row, textvariable=self.ai_training_var, bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT, relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        self.ai_entry.pack(side="left", fill="x", expand=True)
        tk.Button(ai_path_row, text="Browse…", command=self._browse_ai_folder,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="left", padx=6)
        tk.Button(content, text="Training Library Guide", command=self._show_training_info,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=10, pady=6, cursor="hand2").pack(anchor="w", padx=14, pady=(0,8))
        ai_btns = tk.Frame(content, bg=PANEL_DARK)
        ai_btns.pack(fill="x", padx=14, pady=(0,10))
        tk.Button(ai_btns, text="Train Model", command=self._train_ai_model,
                  bg=GPI_HL, fg=GPI_GREEN, relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left")
        tk.Button(ai_btns, text="Export Last Calls…", command=self._export_ai_calls,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left", padx=(10,0))
        footer = tk.Frame(self, bg=PANEL_DARK)
        footer.pack(fill="x", padx=14, pady=(8,14))
        tk.Button(footer, text="Cancel", command=self.destroy,
                  bg="#243B2F" if THEME_MODE=="dark" else "#DFE4DF",
                  fg=TEXT_LIGHT if THEME_MODE=="dark" else "#183024",
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="right")
        tk.Button(footer, text="Apply", command=self._apply,
                  bg=GPI_HL, fg=GPI_GREEN, relief="flat", padx=16, pady=6, cursor="hand2").pack(side="right", padx=6)
    def _browse_tess(self):
        p = filedialog.askopenfilename(title="Select tesseract.exe", filetypes=[("tesseract.exe","tesseract.exe"),("All files","*.*")])
        if not p: return
        self.tesseract_var.set(p)
    def _choose_spcs(self):
        self.master.open_spcs_dialog()
        value = getattr(self.master, "selected_spcs", "")
        self._spcs_label.config(text=value or "None")
    def _browse_ai_folder(self):
        folder = self.master.select_ai_training_folder(parent=self)
        if folder:
            self.ai_training_var.set(folder)

    def _train_ai_model(self):
        value = (self.ai_training_var.get() or "").strip()
        self.master.ai_training_folder_var.set(value)
        self.master.settings["ai_training_dir"] = value
        self.master._save_user_config()
        self.master.train_deed_ai_model(parent=self)

    def _export_ai_calls(self):
        self.master.save_deed_ai_calls(parent=self)

    def _show_training_info(self):
        info = (
            "Training library structure:\n"
            " • Store deed PDF files (*.pdf) directly in the selected folder.\n"
            " • Provide a labeled Excel workbook for each PDF using the same base name\n"
            "   (example.pdf → example_correct.xlsx or example.xlsx).\n"
            " • The Excel file must include a 'Deed_Call_Text' column containing the text\n"
            "   of each call and a 'Start_End' column with start,end character positions\n"
            "   (leave blank to auto-align).\n"
            " • The tool OCRs each PDF to align annotations—set the Tesseract path if\n"
            "   the PDFs are scans.\n"
            " • Only documents with both PDF and labeled calls are used for training."
        )
        messagebox.showinfo("Training Library Requirements", info, parent=self)

    def _apply(self):
        self.on_apply(self.mode_var.get(), self.units_in_var.get(), self.units_out_var.get(),
                      self.bearing_var.get(), self.tesseract_var.get(), self.ai_training_var.get())
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
    STYLE_NAME = "GPI.EditableGrid.Treeview"

    def __init__(self, parent, columns, on_edit_commit, **kwargs):
        style = ttk.Style()
        try:
            fg_normal = TEXT_LIGHT if THEME_MODE == "dark" else "#122016"
            bg_normal = CONSOLE_BG if THEME_MODE == "dark" else "#FFFFFF"
            fg_selected = PANEL_DARK if THEME_MODE == "dark" else "#122016"
            style.configure(
                self.STYLE_NAME,
                foreground=fg_normal,
                background=bg_normal,
                fieldbackground=bg_normal,
                borderwidth=0,
            )
            style.map(
                self.STYLE_NAME,
                background=[("selected", GPI_HL)],
                foreground=[("selected", fg_selected)],
            )
        except Exception:
            pass

        kwargs.setdefault("style", self.STYLE_NAME)
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
        editor_kwargs = {"relief": "flat"}
        if THEME_MODE == "dark":
            editor_kwargs.update(bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT)
        else:
            editor_kwargs.update(bg="#FFFFFF", fg="#122016", insertbackground="#122016")
        self._editor = tk.Entry(self, **editor_kwargs); self._editor.insert(0, value)
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
        self.config_path = Path.home() / ".geo_builder.ini"
        self._config_parser = configparser.ConfigParser()
        self._user_config = {}
        self._load_user_config()
        self.settings = {"theme":THEME_MODE,
                         "units_in":self._user_config.get("units_in","feet"),
                         "units_out":self._user_config.get("units_out","feet"),
                         "bearing_fmt":self._user_config.get("bearing_fmt","dms"),
                         "tesseract_path":self._user_config.get("tesseract_path",""),
                         "spcs_name":self._user_config.get("spcs_name",""),
                         "spcs_epsg":self._coerce_int(self._user_config.get("spcs_epsg")),
                         "origin_easting":self._user_config.get("origin_easting","0.0"),
                         "origin_northing":self._user_config.get("origin_northing","0.0"),
                         "source_epsg":self._user_config.get("source_epsg",""),
                         "apply_pyproj":self._user_config.get("apply_pyproj","false").lower() == "true",
                         "ai_training_dir":self._user_config.get("ai_training_dir", "")}
        self.selected_spcs = self.settings.get("spcs_name") or ""
        self.selected_spcs_epsg = self.settings.get("spcs_epsg")
        self.deed_df = pandas.DataFrame() if pandas else None
        self.deed_pdf_path = None; self.deed_last_saved_excel = None
        self.console = None
        self._log_history = ["Ready."]
        self.grid_row_states = {}
        self._edited_rows = set()
        self.manual_call_entries: List[Dict[str, Any]] = []
        self.origin_easting_var = tk.StringVar(value=self.settings.get("origin_easting","0.0"))
        self.origin_northing_var = tk.StringVar(value=self.settings.get("origin_northing","0.0"))
        self.source_epsg_var = tk.StringVar(value=self.settings.get("source_epsg",""))
        self.apply_pyproj_var = tk.BooleanVar(value=self.settings.get("apply_pyproj", False))
        self.parcel_points_rel: List[Tuple[float,float]] = []
        self.parcel_points_abs: List[Tuple[float,float]] = []
        self.parcel_segments: List[ParcelSegment] = []
        self.points_tree = None
        self.canvas = None
        self.ax = None
        self.figure = None
        self.ai_training_folder_var = tk.StringVar(value=self.settings.get("ai_training_dir", ""))
        self.deed_ai_model = None
        self.deed_ai_model_path = Path(__file__).with_name("deed_ner_model")
        self.deed_ai_last_results: List[Tuple[str, Tuple[int, int]]] = []
        self.ai_output_text = None
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
        extract_row = tk.Frame(parent, bg=PANEL_DARK); extract_row.pack(fill="x", padx=16, pady=(10,6))
        self.btn_extract_text = self._cta_button(extract_row, "Extract Text"); self.btn_extract_text.pack(side="left")
        self.btn_extract_text.configure(command=self.extract_deed_text)
        self._bind_hint(self.btn_extract_text, "Extract deed text into an editable preview")
        self.btn_ai_extract_text = self._cta_button(extract_row, "AI Extract Text"); self.btn_ai_extract_text.pack(side="left", padx=(10,0))
        self.btn_ai_extract_text.configure(command=self.ai_extract_text_and_calls)
        self._bind_hint(self.btn_ai_extract_text, "Extract deed text with OCR and parse calls using the AI model")

        btns = tk.Frame(parent, bg=PANEL_DARK); btns.pack(fill="x", padx=16, pady=(6,6))
        self.btn_add_line = self._secondary_button(btns, "Add Line", self.add_manual_line); self.btn_add_line.pack(side="left")
        self._bind_hint(self.btn_add_line, "Select line text in the deed and add it to the call list")
        self.btn_edit_line = self._secondary_button(btns, "Edit Line", self.edit_manual_line); self.btn_edit_line.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_edit_line, "Adjust the highlighted text for an existing manual line call")
        self.btn_add_curve = self._secondary_button(btns, "Add Curve", self.add_manual_curve); self.btn_add_curve.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_add_curve, "Select curve text in the deed and add it to the call list")
        self.btn_edit_curve = self._secondary_button(btns, "Edit Curve", self.edit_manual_curve); self.btn_edit_curve.pack(side="left", padx=(10,0))
        self._bind_hint(self.btn_edit_curve, "Adjust the highlighted text for an existing manual curve call")
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
        _bind_mousewheel_scroll(self.deed_text)
        self._configure_call_highlight_tags()
        info_msg = ("Extract the deed PDF to populate the text above.\n"
                    "Review and edit as needed before running call extraction from the next tab.")
        tk.Label(parent, text=info_msg, justify="left", bg=PANEL_DARK, fg=TEXT_SOFT, font=("Segoe UI",10)).pack(anchor="w", padx=16, pady=(0,12))

    def select_ai_training_folder(self, parent=None):
        folder = filedialog.askdirectory(parent=parent or self, title="Select Deed AI Training Folder")
        if not folder:
            return None
        self.ai_training_folder_var.set(folder)
        self.settings["ai_training_dir"] = folder
        self._save_user_config()
        self._log(f"AI training folder set → {folder}")
        return folder

    def train_deed_ai_model(self, parent=None):
        if spacy is None:
            messagebox.showerror("spaCy not available", "spaCy is required for AI training and could not be installed automatically.\nInstall manually with: pip install spacy", parent=parent or self)
            return
        folder_path = Path(self.ai_training_folder_var.get() or "")
        if not folder_path.exists():
            messagebox.showwarning("Training folder missing", "Select a training folder that contains deed PDFs and labeled Excel files.", parent=parent or self)
            return
        try:
            dataset = load_deed_training_dataset(folder_path, tesseract_path=self.settings.get("tesseract_path"), logger=self._log)
            if not dataset:
                messagebox.showinfo("No training data", "No labeled deed calls were found in the selected folder.", parent=parent or self)
                return
            model = train_deed_spacy_model(dataset, output_dir=self.deed_ai_model_path, logger=self._log)
            self.deed_ai_model = model
            messagebox.showinfo("Training complete", f"Model trained with {len(dataset)} document(s).", parent=parent or self)
            self._log("Deed AI model trained successfully.")
        except Exception as exc:
            self._log(f"AI training failed: {exc}")
            messagebox.showerror("Training failed", str(exc), parent=parent or self)

    def run_deed_ai_analysis(self):
        if spacy is None:
            messagebox.showerror("spaCy not available", "spaCy is required for deed analysis and could not be installed automatically.\nInstall manually with: pip install spacy", parent=self)
            return
        model = self._get_or_load_deed_ai_model()
        if model is None:
            messagebox.showwarning("Model unavailable", "Train the AI model or install spaCy's en_core_web_sm package.", parent=self)
            return
        pdf_path = filedialog.askopenfilename(parent=self, title="Select deed PDF for AI analysis", filetypes=[("PDF", "*.pdf"), ("All files", "*.*")])
        if not pdf_path:
            return
        pdf_path_obj = Path(pdf_path)
        try:
            text = ocr_pdf_with_pytesseract(pdf_path_obj, tesseract_path=self.settings.get("tesseract_path"), logger=self._log)
            if not text:
                messagebox.showwarning("No OCR text", "OCR did not return any text for this PDF.", parent=self)
                return
            calls = extract_deed_calls_with_model(model, text)
            self.deed_ai_last_results = calls
            self._update_ai_output(calls)
            if calls:
                self._log(f"AI extracted {len(calls)} deed call(s) from {pdf_path_obj.name}.")
            else:
                self._log(f"AI found no deed calls in {pdf_path_obj.name}.")
        except Exception as exc:
            self._log(f"AI analysis failed: {exc}")
            messagebox.showerror("Analysis failed", str(exc), parent=self)

    def save_deed_ai_calls(self, parent=None):
        if not self.deed_ai_last_results:
            messagebox.showinfo("No results", "Run an AI analysis before saving calls.", parent=parent or self)
            return
        save_path = filedialog.asksaveasfilename(parent=parent or self, title="Save AI deed calls", defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not save_path:
            return
        try:
            with open(save_path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh)
                writer.writerow(["Deed_Call_Text", "Start_Index", "End_Index"])
                for text_value, (start, end) in self.deed_ai_last_results:
                    writer.writerow([text_value, start, end])
            self._log(f"AI deed calls saved → {save_path}")
            messagebox.showinfo("Saved", f"AI deed calls exported to:\n{save_path}", parent=parent or self)
        except Exception as exc:
            self._log(f"Failed to save AI calls: {exc}")
            messagebox.showerror("Save failed", str(exc), parent=parent or self)

    def _update_ai_output(self, calls: List[Tuple[str, Tuple[int, int]]]):
        if not getattr(self, "ai_output_text", None):
            return
        self.ai_output_text.configure(state="normal")
        self.ai_output_text.delete("1.0", "end")
        if not calls:
            self.ai_output_text.insert("end", "No deed calls extracted yet. Train and analyze to see results.\n")
        else:
            for idx, (text_value, (start, end)) in enumerate(calls, start=1):
                self.ai_output_text.insert("end", f"{idx}. {text_value}\n    (start: {start}, end: {end})\n\n")
        self.ai_output_text.configure(state="disabled")

    def _get_or_load_deed_ai_model(self):
        if spacy is None:
            return None
        if self.deed_ai_model is not None:
            return self.deed_ai_model
        if self.deed_ai_model_path.exists():
            try:
                self.deed_ai_model = spacy.load(self.deed_ai_model_path)
                self._log("Loaded saved deed AI model.")
                return self.deed_ai_model
            except Exception as exc:
                self._log(f"Failed to load saved model: {exc}")
        if ensure_spacy_model():
            try:
                self.deed_ai_model = spacy.load("en_core_web_sm")
                self._log("Loaded spaCy en_core_web_sm model as fallback.")
            except Exception as exc:
                self._log(f"Failed to load en_core_web_sm: {exc}")
                self.deed_ai_model = None
        if self.deed_ai_model is None:
            self.deed_ai_model = spacy.blank("en")
            if self.deed_ai_model is not None:
                self._log("Initialized blank spaCy English model. Training is recommended before analysis.")
        if self.deed_ai_model is not None:
            if "ner" not in self.deed_ai_model.pipe_names:
                ner = self.deed_ai_model.add_pipe("ner")
            else:
                ner = self.deed_ai_model.get_pipe("ner")
            if "DEED_CALL" not in getattr(ner, "labels", []):
                try:
                    ner.add_label("DEED_CALL")
                except Exception:
                    pass
        return self.deed_ai_model

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
        _bind_mousewheel_scroll(self.grid)
        self._configure_grid_tags()
        if pandas is None: self._log("pandas not installed — PDF parsing and Excel saving will be disabled until installed (pip install pandas).")
        if openpyxl is None: self._log("openpyxl not installed — saving Excel will be disabled (pip install openpyxl).")
        if pdfplumber is None and fitz is None: self._log("pdfplumber / PyMuPDF not installed — text extraction may fail (pip install pdfplumber pymupdf).")
        if pytesseract is None: self._log("pytesseract not installed — OCR fallback disabled (pip install pytesseract).")
        self._refresh_grid_from_df()

        geom_wrapper = tk.Frame(parent, bg=PANEL_DARK)
        geom_wrapper.pack(fill="both", expand=True, padx=16, pady=(4,12))

        control_frame = tk.Frame(geom_wrapper, bg=PANEL_DARK)
        control_frame.pack(side="left", fill="y")

        pob_lbl = tk.Label(control_frame, text="Point of Beginning (absolute)",
                           bg=PANEL_DARK, fg=TEXT_LIGHT, font=("Segoe UI",10,"bold"))
        pob_lbl.pack(anchor="w", pady=(0,4))
        pob_frame = tk.Frame(control_frame, bg=PANEL_DARK)
        pob_frame.pack(anchor="w", fill="x")
        tk.Label(pob_frame, text="Easting / X", bg=PANEL_DARK, fg=TEXT_SOFT,
                 font=("Segoe UI",9)).grid(row=0, column=0, sticky="w")
        e_entry = tk.Entry(pob_frame, textvariable=self.origin_easting_var,
                           bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                           relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER, width=14)
        e_entry.grid(row=1, column=0, sticky="we", pady=(0,4))
        self._bind_hint(e_entry, "Absolute easting/X for the parcel point of beginning.")

        tk.Label(pob_frame, text="Northing / Y", bg=PANEL_DARK, fg=TEXT_SOFT,
                 font=("Segoe UI",9)).grid(row=0, column=1, sticky="w", padx=(12,0))
        n_entry = tk.Entry(pob_frame, textvariable=self.origin_northing_var,
                           bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                           relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER, width=14)
        n_entry.grid(row=1, column=1, sticky="we", padx=(12,0), pady=(0,4))
        self._bind_hint(n_entry, "Absolute northing/Y for the parcel point of beginning.")

        pob_frame.columnconfigure(0, weight=1)
        pob_frame.columnconfigure(1, weight=1)

        spcs_frame = tk.Frame(control_frame, bg=PANEL_DARK)
        spcs_frame.pack(anchor="w", fill="x", pady=(8,4))
        tk.Label(spcs_frame, text="State Plane Coordinate System", bg=PANEL_DARK,
                 fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        tk.Label(spcs_frame, text="Selected:", bg=PANEL_DARK, fg=TEXT_SOFT,
                 font=("Segoe UI",9)).grid(row=1, column=0, sticky="w")
        self.spcs_value_lbl = tk.Label(spcs_frame, text=self.selected_spcs or "None", bg=PANEL_DARK,
                                       fg=TEXT_LIGHT, font=("Segoe UI",9), wraplength=220, justify="left")
        self.spcs_value_lbl.grid(row=1, column=1, sticky="w", padx=(6,0))
        choose_btn = self._secondary_button(spcs_frame, "Choose…", self.open_spcs_dialog)
        choose_btn.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4,0))
        self._bind_hint(choose_btn, "Pick a State Plane zone for DXF metadata.")

        tk.Label(spcs_frame, text="Source EPSG (optional)", bg=PANEL_DARK, fg=TEXT_SOFT,
                 font=("Segoe UI",9)).grid(row=3, column=0, sticky="w", pady=(8,0))
        epsg_entry = tk.Entry(spcs_frame, textvariable=self.source_epsg_var, width=16,
                              bg=CONSOLE_BG, fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT,
                              relief="flat", highlightthickness=1, highlightbackground=PANEL_BORDER)
        epsg_entry.grid(row=3, column=1, sticky="w", padx=(6,0), pady=(8,0))
        self._bind_hint(epsg_entry, "EPSG of current working coordinates for optional pyproj transform.")

        transform_chk = tk.Checkbutton(spcs_frame, text="Transform with pyproj when exporting",
                                       variable=self.apply_pyproj_var,
                                       bg=PANEL_DARK, fg=TEXT_LIGHT, selectcolor=PANEL_DARK,
                                       activebackground=PANEL_DARK, activeforeground=TEXT_LIGHT,
                                       state="normal" if HAVE_PYPROJ else "disabled")
        transform_chk.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6,0))
        if not HAVE_PYPROJ:
            self._bind_hint(transform_chk, "Install pyproj to enable coordinate transformations.")

        btn_frame = tk.Frame(control_frame, bg=PANEL_DARK)
        btn_frame.pack(fill="x", pady=(10,6))
        process_btn = self._cta_button(btn_frame, "Process Geometry")
        process_btn.pack(side="left")
        process_btn.configure(command=self.process_deed_geometry)
        self._bind_hint(process_btn, "Compute parcel coordinates from the parsed calls.")

        export_xml_btn = self._secondary_button(btn_frame, "Export LandXML", self.export_landxml)
        export_xml_btn.pack(side="left", padx=(10,0))
        self._bind_hint(export_xml_btn, "Write a LandXML parcel using the offset points.")

        export_dxf_btn = self._secondary_button(btn_frame, "Export DXF", self.export_dxf)
        export_dxf_btn.pack(side="left", padx=(10,0))
        hint = "Create a DXF parcel. Install ezdxf for this export." if not HAVE_EZDXF else "Create a DXF parcel file."
        self._bind_hint(export_dxf_btn, hint)
        if not HAVE_EZDXF:
            export_dxf_btn.configure(state="disabled")

        tk.Label(control_frame, text="Computed Parcel Points", bg=PANEL_DARK,
                 fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", pady=(6,2))
        points_container = tk.Frame(control_frame, bg=PANEL_DARK)
        points_container.pack(fill="both", expand=True)
        columns = ("#", "ΔE", "ΔN", "E", "N")
        self.points_tree = ttk.Treeview(points_container, columns=columns, show="headings", height=10)
        headings = ["#", "ΔE", "ΔN", "Easting", "Northing"]
        for col, title in zip(columns, headings):
            anchor = "center" if col == "#" else "e"
            self.points_tree.heading(col, text=title)
            self.points_tree.column(col, width=90, anchor=anchor, stretch=True)
        self.points_tree.pack(side="left", fill="both", expand=True)
        pts_scroll = tk.Scrollbar(points_container, orient="vertical", command=self.points_tree.yview)
        self.points_tree.configure(yscrollcommand=pts_scroll.set)
        pts_scroll.pack(side="right", fill="y")
        _bind_mousewheel_scroll(self.points_tree, points_container)

        preview_frame = tk.Frame(geom_wrapper, bg=PANEL_DARK)
        preview_frame.pack(side="left", fill="both", expand=True, padx=(18,0))
        tk.Label(preview_frame, text="Parcel Preview", bg=PANEL_DARK,
                 fg=TEXT_LIGHT, font=("Segoe UI",10,"bold")).pack(anchor="w", pady=(0,4))
        if HAVE_MPL:
            self.figure = Figure(figsize=(4.5,3.2), dpi=100)
            self.ax = self.figure.add_subplot(111)
            self.ax.set_title("Plan View")
            self.ax.set_aspect("equal", adjustable="datalim")
            self.ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)
            self.canvas = FigureCanvasTkAgg(self.figure, master=preview_frame)
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            msg = "Install matplotlib for live parcel previews."
            tk.Label(preview_frame, text=msg, bg=PANEL_DARK, fg=TEXT_SOFT,
                     font=("Segoe UI",9), justify="left", wraplength=300).pack(fill="both", expand=True)
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
            self.manual_call_entries = []
        if getattr(self, "pb_deed", None): self.pb_deed["value"] = 100
        self.deed_pdf_path = p
        self._log("Deed text ready for QC. Review/edit before running call extraction.")
        if txt and txt.strip():
            self.highlight_calls_preview(quiet=True)


    def ai_extract_text_and_calls(self):
        if spacy is None:
            messagebox.showerror("spaCy not available", "spaCy is required for AI extraction and could not be installed automatically.\nInstall manually with: pip install spacy")
            return
        p = Path(self.pdf_var.get() or "")
        if not p or not p.exists() or p.suffix.lower() != ".pdf":
            messagebox.showerror("Missing or invalid PDF", "Please select a valid .pdf deed file before running AI extraction.")
            return
        model = self._get_or_load_deed_ai_model()
        if model is None:
            messagebox.showwarning("Model unavailable", "Train the AI model from Settings before running AI extraction.")
            return
        self._log(f"AI extract & parse → {p}")
        progress = getattr(self, "pb_deed", None)
        if progress:
            progress["value"] = 0
            progress.update_idletasks()

        def _update_progress(done_pages: int, total_pages: int, stage: str = "ocr"):
            if not progress:
                return
            total_pages = max(total_pages, 1)
            if stage == "ocr":
                ratio = min(max(done_pages, 0), total_pages) / total_pages
                progress["value"] = 10 + ratio * 60
            else:
                progress["value"] = max(progress["value"], 80)
            progress.update_idletasks()
        tess_path = self.settings.get("tesseract_path")
        if tess_path:
            try_set_tesseract_cmd(tess_path)
        try:
            text_value = ocr_pdf_with_pytesseract(
                p,
                tesseract_path=tess_path,
                logger=self._log,
                progress_callback=_update_progress,
            ) or ""
        except Exception as exc:
            if progress:
                progress["value"] = 0
                progress.update_idletasks()
            messagebox.showerror("AI extraction failed", str(exc))
            self._log(f"AI extraction failed: {exc}")
            return
        if not text_value.strip():
            if progress:
                progress["value"] = 0
                progress.update_idletasks()
            messagebox.showwarning("No OCR text", "The AI extraction did not return any text for this PDF.")
            self._log("AI extraction produced no text.")
            return
        if getattr(self, "deed_text", None):
            self.deed_text.delete("1.0", "end")
            self.deed_text.insert("1.0", text_value)
        if progress:
            progress["value"] = max(progress["value"], 75)
            progress.update_idletasks()
        self.deed_pdf_path = p
        self.manual_call_entries = []
        self.deed_ai_last_results = []
        try:
            calls = extract_deed_calls_with_model(model, text_value)
            self.deed_ai_last_results = calls
        except Exception as exc:
            if progress:
                progress["value"] = 0
                progress.update_idletasks()
            messagebox.showerror("AI parsing failed", str(exc))
            self._log(f"AI parsing failed: {exc}")
            return
        assumed_unit = self.settings.get("units_in", "feet")
        ai_entries = []
        for snippet, span in calls or []:
            start, end = span
            cleaned = clean_text_for_parsing(snippet)
            try:
                parsed = _parse_deed_text_entries(cleaned, assumed_unit)
            except Exception:
                parsed = []
            added = False
            for _, _, data in parsed:
                if not isinstance(data, dict):
                    continue
                call_type = str(data.get("Type", "")).strip().title()
                if call_type not in {"Line", "Curve"}:
                    continue
                normalized = dict(data)
                if call_type == "Line" and normalized.get("Distance") not in (None, "") and not normalized.get("DistanceUnit"):
                    normalized["DistanceUnit"] = assumed_unit
                if call_type == "Curve":
                    for unit_key, value_key in (("RadiusUnit", "Radius"), ("ArcUnit", "Arc Length"), ("ChordUnit", "Chord Length")):
                        if normalized.get(value_key) not in (None, "") and not normalized.get(unit_key):
                            normalized[unit_key] = assumed_unit
                ai_entries.append({"start": max(0, start), "end": min(len(text_value), end), "type": call_type, "data": normalized})
                added = True
                break
            if not added and snippet.strip():
                self._log(f"AI call could not be categorized: {snippet[:80]}…")
        if ai_entries:
            ai_entries.sort(key=lambda item: item.get("start", 0))
            self.manual_call_entries = ai_entries
        if getattr(self, "deed_text", None):
            self.deed_text.tag_remove("call_line", "1.0", "end")
            self.deed_text.tag_remove("call_curve", "1.0", "end")
        self._apply_manual_call_tags()
        self.highlight_calls_preview(quiet=True)
        if pandas is not None:
            try:
                self.extract_calls_from_text()
            except Exception as exc:
                self._log(f"AI call parsing error: {exc}")
        else:
            self._log("pandas not available; skipping call parsing after AI extraction.")
        msg = f"AI extracted {len(ai_entries)} call(s)." if ai_entries else "AI extraction finished with no calls detected."
        self._log(msg)
        if not ai_entries:
            messagebox.showinfo("No calls detected", "The AI model did not detect any deed calls in this PDF.")
        if progress:
            progress["value"] = 100
            progress.update_idletasks()

    def clear_deed_text(self):
        if getattr(self, "deed_text", None):
            self.deed_text.delete("1.0", "end")
            self.deed_text.tag_remove("call_line", "1.0", "end")
            self.deed_text.tag_remove("call_curve", "1.0", "end")
        self.manual_call_entries = []
        if getattr(self, "pb_deed", None): self.pb_deed["value"] = 0
        if pandas is not None:
            self.deed_df = pandas.DataFrame()
        else:
            self.deed_df = None
        self._edited_rows = set()
        self._refresh_grid_from_df()

    def _parse_deed_text_with_manual_entries(self, deed_text: str, assumed_unit: str):
        if pandas is None:
            raise RuntimeError("pandas is required to build the parsed table. Please install:\n  pip install pandas")
        columns = [
            "Type",
            "Bearing",
            "Distance",
            "DistanceUnit",
            "Radius",
            "RadiusUnit",
            "Arc Length",
            "ArcUnit",
            "Chord Length",
            "ChordUnit",
            "Chord Bearing",
        ]
        cleaned_text, mapping = _clean_text_for_parsing_with_map(deed_text)
        ordered_entries = _parse_deed_text_entries(cleaned_text, assumed_unit)
        combined: List[Tuple[int, Dict[str, Any], bool]] = []
        for start, _, data in ordered_entries:
            if not isinstance(data, dict):
                continue
            orig_start = mapping[start] if start < len(mapping) else len(deed_text)
            combined.append((orig_start, dict(data), False))
        manual_count = 0
        for entry in self.manual_call_entries:
            start = entry.get("start")
            end = entry.get("end")
            if not isinstance(start, int) or start < 0:
                continue
            call_type = str(entry.get("type", "") or "").strip().title()
            stored_data = entry.get("data") if isinstance(entry.get("data"), dict) else None
            if not call_type and isinstance(stored_data, dict):
                call_type = str(stored_data.get("Type", "") or "").strip().title()
            snippet = ""
            if isinstance(end, int) and start < end <= len(deed_text):
                snippet = deed_text[start:end]
            elif start < len(deed_text):
                snippet = deed_text[start: min(start + 400, len(deed_text))]
            refreshed = stored_data
            if snippet.strip() and call_type:
                try:
                    refreshed = self._analyze_manual_call(snippet, call_type)
                except Exception:
                    pass
            if not isinstance(refreshed, dict):
                continue
            entry["data"] = refreshed
            combined.append((start, dict(refreshed), True))
        if not combined:
            return pandas.DataFrame(columns=columns), manual_count
        combined.sort(key=lambda item: (item[0], 1 if item[2] else 0))
        rows: List[Dict[str, Any]] = []
        seen_keys = set()
        for position, data, is_manual in combined:
            if not isinstance(data, dict):
                continue
            row = {
                "Type": data.get("Type"),
                "Bearing": data.get("Bearing"),
                "Distance": data.get("Distance"),
                "DistanceUnit": data.get("DistanceUnit"),
                "Radius": data.get("Radius"),
                "RadiusUnit": data.get("RadiusUnit"),
                "Arc Length": data.get("Arc Length"),
                "ArcUnit": data.get("ArcUnit"),
                "Chord Length": data.get("Chord Length"),
                "ChordUnit": data.get("ChordUnit"),
                "Chord Bearing": data.get("Chord Bearing"),
            }
            typ_lower = str(row.get("Type") or "").strip().lower()
            if typ_lower == "line" and row.get("Distance") not in (None, "") and not row.get("DistanceUnit"):
                row["DistanceUnit"] = assumed_unit
            if typ_lower == "curve":
                for unit_key, value_key in (("RadiusUnit", "Radius"), ("ArcUnit", "Arc Length"), ("ChordUnit", "Chord Length")):
                    if row.get(value_key) not in (None, "") and not row.get(unit_key):
                        row[unit_key] = assumed_unit
            key = (
                position,
                typ_lower,
                row.get("Bearing"),
                row.get("Distance"),
                row.get("Radius"),
                row.get("Arc Length"),
                row.get("Chord Length"),
                row.get("Chord Bearing"),
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)
            rows.append(row)
            if is_manual:
                manual_count += 1
        df = pandas.DataFrame(rows, columns=columns)
        return df, manual_count

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
            df_parsed, manual_count = self._parse_deed_text_with_manual_entries(deed_text, deed_units_default)
            if manual_count:
                logger(f"Parsed rows: {len(df_parsed)} (including {manual_count} manual call(s))")
            else:
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
        self._apply_manual_call_tags()
        self._log(f"Call preview: highlighted {len(spans)} potential call(s).")

    def add_manual_line(self):
        self._add_manual_call("Line")

    def add_manual_curve(self):
        self._add_manual_call("Curve")

    def edit_manual_line(self):
        self._edit_manual_call("Line")

    def edit_manual_curve(self):
        self._edit_manual_call("Curve")

    def _add_manual_call(self, call_type: str):
        if not getattr(self, "deed_text", None):
            return
        try:
            start_index = self.deed_text.index("sel.first")
            end_index = self.deed_text.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight the deed text for the call before adding it.")
            return
        if start_index == end_index:
            messagebox.showinfo("Select text", "Highlight the deed text for the call before adding it.")
            return
        snippet = self.deed_text.get(start_index, end_index)
        if not snippet.strip():
            messagebox.showinfo("Empty selection", "The selected text was empty. Highlight the call text and try again.")
            return
        try:
            call_data = self._analyze_manual_call(snippet, call_type)
        except ValueError as exc:
            messagebox.showerror("Call not recognized", str(exc))
            return
        start_chars = int(self.deed_text.count("1.0", start_index, "chars")[0])
        end_chars = int(self.deed_text.count("1.0", end_index, "chars")[0])
        if end_chars <= start_chars:
            end_chars = start_chars + len(snippet)
        tag = "call_curve" if call_type.lower() == "curve" else "call_line"
        try:
            self.deed_text.tag_add(tag, start_index, end_index)
        except tk.TclError:
            pass
        self.manual_call_entries = [
            entry for entry in self.manual_call_entries
            if not (self._spans_overlap(entry.get("start", 0), entry.get("end", 0), start_chars, end_chars))
        ]
        self.manual_call_entries.append({
            "start": start_chars,
            "end": end_chars,
            "type": call_type,
            "data": call_data,
        })
        self.manual_call_entries.sort(key=lambda item: item.get("start", 0))
        self._log(f"Added manual {call_type.lower()} call from highlighted text.")

    def _edit_manual_call(self, call_type: str):
        if not getattr(self, "deed_text", None):
            return
        try:
            start_index = self.deed_text.index("sel.first")
            end_index = self.deed_text.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight the deed text for the call before editing it.")
            return
        if start_index == end_index:
            messagebox.showinfo("Select text", "Highlight the deed text for the call before editing it.")
            return
        snippet = self.deed_text.get(start_index, end_index)
        if not snippet.strip():
            messagebox.showinfo("Empty selection", "The selected text was empty. Highlight the call text and try again.")
            return
        try:
            call_data = self._analyze_manual_call(snippet, call_type)
        except ValueError as exc:
            messagebox.showerror("Call not recognized", str(exc))
            return
        start_chars = int(self.deed_text.count("1.0", start_index, "chars")[0])
        end_chars = int(self.deed_text.count("1.0", end_index, "chars")[0])
        if end_chars <= start_chars:
            end_chars = start_chars + len(snippet)
        tag = "call_curve" if call_type.lower() == "curve" else "call_line"
        target_entry = None
        for entry in self.manual_call_entries:
            if str(entry.get("type", "")).lower() != call_type.lower():
                continue
            entry_start = entry.get("start", 0)
            entry_end = entry.get("end", 0)
            if self._spans_overlap(entry_start, entry_end, start_chars, end_chars) or (
                entry_start <= start_chars <= entry_end
            ) or (
                entry_start <= end_chars <= entry_end
            ):
                target_entry = entry
                break
        if not target_entry:
            messagebox.showinfo(
                "No manual call",
                f"No manual {call_type.lower()} call overlaps the selection. Add one first, then edit it.",
            )
            return
        try:
            self.deed_text.tag_remove(tag, f"1.0+{target_entry.get('start', 0)}c", f"1.0+{target_entry.get('end', 0)}c")
        except tk.TclError:
            pass
        target_entry["start"] = start_chars
        target_entry["end"] = end_chars
        target_entry["data"] = call_data
        try:
            self.deed_text.tag_add(tag, start_index, end_index)
        except tk.TclError:
            pass
        to_remove = []
        for entry in self.manual_call_entries:
            if entry is target_entry:
                continue
            if str(entry.get("type", "")).lower() != call_type.lower():
                continue
            if self._spans_overlap(entry.get("start", 0), entry.get("end", 0), start_chars, end_chars):
                to_remove.append(entry)
        for entry in to_remove:
            try:
                self.deed_text.tag_remove(tag, f"1.0+{entry.get('start', 0)}c", f"1.0+{entry.get('end', 0)}c")
            except tk.TclError:
                pass
        self.manual_call_entries = [entry for entry in self.manual_call_entries if entry not in to_remove]
        if target_entry not in self.manual_call_entries:
            self.manual_call_entries.append(target_entry)
        self.manual_call_entries.sort(key=lambda item: item.get("start", 0))
        self._log(f"Updated manual {call_type.lower()} call from highlighted text.")

    def _analyze_manual_call(self, snippet: str, call_type: str) -> Dict[str, Any]:
        assumed_unit = self.settings.get("units_in", "feet")
        cleaned = clean_text_for_parsing(snippet)
        if not cleaned.strip():
            raise ValueError("The selected text did not contain any call information.")
        entries = _parse_deed_text_entries(cleaned, assumed_unit)
        expected_type = call_type.lower()
        if not entries:
            raise ValueError("No call could be parsed from the selected text.")
        for _, _, data in entries:
            typ = str(data.get("Type", "")).strip().lower()
            if typ == expected_type:
                normalized = dict(data)
                if expected_type == "line" and normalized.get("Distance") not in (None, "") and not normalized.get("DistanceUnit"):
                    normalized["DistanceUnit"] = assumed_unit
                if expected_type == "curve":
                    for unit_key, value_key in (("RadiusUnit", "Radius"), ("ArcUnit", "Arc Length"), ("ChordUnit", "Chord Length")):
                        if normalized.get(value_key) not in (None, "") and not normalized.get(unit_key):
                            normalized[unit_key] = assumed_unit
                return normalized
        available_types = {str(d.get("Type", "")).strip() for _, _, d in entries if d}
        if available_types:
            parsed_desc = ", ".join(sorted(t for t in available_types if t))
            raise ValueError(f"The selected text matched {parsed_desc} detail(s) but not a {call_type.lower()} call.")
        raise ValueError("The selected text could not be parsed as the requested call type.")

    @staticmethod
    def _spans_overlap(a_start: int, a_end: int, b_start: int, b_end: int) -> bool:
        return max(a_start, b_start) < min(a_end, b_end)

    def _apply_manual_call_tags(self):
        if not getattr(self, "deed_text", None):
            return
        for entry in self.manual_call_entries:
            start = entry.get("start")
            end = entry.get("end")
            typ = str(entry.get("type", "")).lower()
            if start is None or end is None or start >= end:
                continue
            tag = "call_curve" if typ == "curve" else "call_line"
            try:
                self.deed_text.tag_add(tag, f"1.0+{start}c", f"1.0+{end}c")
            except tk.TclError:
                continue

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

    # ---------------------- Parcel geometry ----------------------
    def process_deed_geometry(self):
        if self.deed_df is None or self.deed_df.empty:
            messagebox.showerror("No courses", "Run call extraction before processing geometry.")
            return
        try:
            origin_x = float(self.origin_easting_var.get() or 0.0)
            origin_y = float(self.origin_northing_var.get() or 0.0)
        except ValueError:
            messagebox.showerror("Invalid origin", "Origin coordinates must be numeric values.")
            return
        try:
            calls = self._build_calls_from_df(self.deed_df)
        except ValueError as exc:
            messagebox.showerror("Processing error", str(exc))
            return
        if not calls:
            messagebox.showwarning("No geometry", "No line or curve calls found to process.")
            return
        try:
            result = self._compute_geometry(calls, (origin_x, origin_y))
        except ValueError as exc:
            messagebox.showerror("Geometry error", str(exc))
            return
        self.parcel_points_rel = result.points_rel
        self.parcel_points_abs = result.points_abs
        self.parcel_segments = result.segments
        closed = self._ensure_closed()
        self._update_points_tree()
        self._update_preview()
        self._save_user_config()
        msg = "Parcel geometry processed."
        if not closed:
            msg += " Parcel was forced closed with a final segment."
        self._log(msg)

    def _build_calls_from_df(self, df):
        calls = []
        for _, row in df.iterrows():
            typ = str(row.get("Type", "")).strip().lower()
            if typ not in ("line", "curve", "arc"):
                continue
            bearing = row.get("Bearing")
            distance = self._to_float(row.get("Distance (ft)") or row.get("Distance"))
            radius = self._to_float(row.get("Radius (ft)") or row.get("Radius"))
            arc_length = self._to_float(row.get("Arc Length (ft)") or row.get("Arc Length"))
            chord_length = self._to_float(row.get("Chord Length (ft)") or row.get("Chord Length"))
            chord_bearing = row.get("Chord Bearing")
            rotation = row.get("Rotation") or row.get("Curve Rotation") or row.get("Turn")
            calls.append(ParsedCall(
                type="curve" if typ == "arc" else typ,
                bearing=str(bearing).strip() if bearing not in (None, "") else None,
                distance=distance,
                radius=radius,
                arc_length=arc_length,
                chord_length=chord_length,
                chord_bearing=str(chord_bearing).strip() if chord_bearing not in (None, "") else None,
                rotation=str(rotation).strip() if rotation not in (None, "") else None,
            ))
        return calls

    def _compute_geometry(self, calls: Sequence[ParsedCall], origin: Tuple[float, float]) -> ProcessResult:
        points_rel: List[Tuple[float, float]] = [(0.0, 0.0)]
        points_abs: List[Tuple[float, float]] = [origin]
        segments: List[ParcelSegment] = []
        current_dir = None
        bearing_fmt = self.settings.get("bearing_fmt", "dms")

        for call in calls:
            start_rel = points_rel[-1]
            start_abs = points_abs[-1]
            if call.type == "line":
                if call.distance is None:
                    raise ValueError("Line call missing distance.")
                if call.bearing:
                    current_dir = parse_bearing_to_east_ccw_radians(call.bearing, bearing_fmt)
                elif current_dir is None:
                    current_dir = 0.0
                dx = math.cos(current_dir) * call.distance
                dy = math.sin(current_dir) * call.distance
                end_rel = (start_rel[0] + dx, start_rel[1] + dy)
                end_abs = (start_abs[0] + dx, start_abs[1] + dy)
                points_rel.append(end_rel)
                points_abs.append(end_abs)
                segments.append(ParcelSegment(
                    type="line",
                    start=start_abs,
                    end=end_abs,
                    start_rel=start_rel,
                    end_rel=end_rel,
                    bearing=call.bearing,
                    distance=call.distance,
                ))
            elif call.type in ("curve", "arc"):
                if call.radius is None and call.chord_length is None:
                    raise ValueError("Curve call requires a radius or chord length.")
                rotation = (call.rotation or "CCW").upper()
                if rotation in {"LEFT", "CCW"}:
                    rotation = "CCW"
                elif rotation in {"RIGHT", "CW"}:
                    rotation = "CW"
                else:
                    rotation = "CCW"
                radius = call.radius
                chord = call.chord_length
                if chord is None and call.arc_length and radius:
                    chord = 2 * radius * math.sin((call.arc_length / radius) / 2.0)
                if radius is None and chord is not None:
                    radius = chord / (2 * math.sin(math.radians(45))) if chord else None
                if radius is None or chord is None:
                    raise ValueError("Unable to determine radius/chord for curve call.")
                if chord / 2.0 > radius:
                    raise ValueError("Curve chord exceeds diameter for given radius.")
                if call.arc_length and radius:
                    delta = call.arc_length / radius
                else:
                    delta = 2 * math.asin(max(min(chord / (2 * radius), 1.0), -1.0))
                if call.chord_bearing:
                    chord_angle = parse_bearing_to_east_ccw_radians(call.chord_bearing, bearing_fmt)
                    current_dir = chord_angle + (delta / 2.0 if rotation == "CCW" else -delta / 2.0)
                elif current_dir is not None:
                    turn = delta if rotation == "CCW" else -delta
                    chord_angle = current_dir + turn / 2.0
                    current_dir = current_dir + turn
                else:
                    chord_angle = 0.0
                    current_dir = chord_angle + (delta / 2.0 if rotation == "CCW" else -delta / 2.0)
                dx = math.cos(chord_angle) * chord
                dy = math.sin(chord_angle) * chord
                end_rel = (start_rel[0] + dx, start_rel[1] + dy)
                end_abs = (start_abs[0] + dx, start_abs[1] + dy)
                center = self._compute_arc_center(start_abs, end_abs, radius, rotation)
                bulge = math.tan(delta / 4.0)
                if rotation == "CW":
                    bulge *= -1
                segments.append(ParcelSegment(
                    type="arc",
                    start=start_abs,
                    end=end_abs,
                    start_rel=start_rel,
                    end_rel=end_rel,
                    radius=radius,
                    rotation=rotation,
                    delta=math.degrees(delta),
                    bulge=bulge,
                    center=center,
                ))
                points_rel.append(end_rel)
                points_abs.append(end_abs)
            else:
                continue
        return ProcessResult(points_rel, points_abs, segments)

    def _compute_arc_center(self, start, end, radius, rotation):
        dx = end[0] - start[0]
        dy = end[1] - start[1]
        chord = math.hypot(dx, dy)
        if chord == 0 or radius is None or radius <= 0:
            return None
        midpoint = ((start[0] + end[0]) / 2.0, (start[1] + end[1]) / 2.0)
        try:
            height = math.sqrt(max(radius ** 2 - (chord / 2.0) ** 2, 0.0))
        except ValueError:
            return None
        ux, uy = dx / chord, dy / chord
        left_normal = (-uy, ux)
        sign = 1.0 if rotation == "CCW" else -1.0
        cx = midpoint[0] + sign * left_normal[0] * height
        cy = midpoint[1] + sign * left_normal[1] * height
        return (cx, cy)

    def _ensure_closed(self):
        if not self.parcel_points_abs:
            return True
        start = self.parcel_points_abs[0]
        end = self.parcel_points_abs[-1]
        if math.hypot(end[0] - start[0], end[1] - start[1]) <= 1e-4:
            return True
        self.parcel_points_rel.append(self.parcel_points_rel[0])
        self.parcel_points_abs.append(start)
        seg = ParcelSegment(
            type="line",
            start=self.parcel_points_abs[-2],
            end=start,
            start_rel=self.parcel_points_rel[-2],
            end_rel=self.parcel_points_rel[0],
            distance=math.hypot(start[0] - self.parcel_points_abs[-2][0], start[1] - self.parcel_points_abs[-2][1]),
        )
        self.parcel_segments.append(seg)
        return False

    def _polygon_is_closed(self, pts, tol=1e-4):
        if len(pts) < 3:
            return False
        start = pts[0]
        end = pts[-1]
        return math.hypot(end[0] - start[0], end[1] - start[1]) <= tol

    def _update_points_tree(self):
        if not self.points_tree:
            return
        self.points_tree.delete(*self.points_tree.get_children())
        for idx, (rel, abs_pt) in enumerate(zip(self.parcel_points_rel, self.parcel_points_abs), start=1):
            vals = (
                idx,
                f"{rel[0]:.3f}",
                f"{rel[1]:.3f}",
                f"{abs_pt[0]:.3f}",
                f"{abs_pt[1]:.3f}",
            )
            self.points_tree.insert("", "end", values=vals)

    def _update_preview(self):
        if not HAVE_MPL or not self.canvas or not self.ax:
            return
        self.ax.clear()
        self.ax.set_title("Plan View")
        self.ax.set_aspect("equal", adjustable="datalim")
        self.ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)
        if len(self.parcel_points_abs) >= 2:
            xs = [p[0] for p in self.parcel_points_abs]
            ys = [p[1] for p in self.parcel_points_abs]
            self.ax.plot(xs, ys, marker="o")
        self.canvas.draw_idle()

    def _prepare_export_geometry(self):
        if not self.parcel_points_abs:
            return [], []
        pts = list(self.parcel_points_abs)
        transformer = None
        if self.apply_pyproj_var.get() and HAVE_PYPROJ and self.selected_spcs_epsg and self.source_epsg_var.get():
            try:
                src_epsg = int(self.source_epsg_var.get())
                transformer = pyproj.Transformer.from_crs(f"EPSG:{src_epsg}", f"EPSG:{self.selected_spcs_epsg}", always_xy=True)
                pts = [transformer.transform(x, y) for x, y in pts]
            except Exception as exc:
                transformer = None
                messagebox.showwarning("Transformation failed", str(exc), parent=self)
        export_segments = []
        for idx, seg in enumerate(self.parcel_segments):
            if idx + 1 >= len(pts):
                break
            start = pts[idx]
            end = pts[idx + 1]
            center = seg.center
            if transformer and seg.center:
                try:
                    center = transformer.transform(seg.center[0], seg.center[1])
                except Exception:
                    center = seg.center
            export_segments.append(ParcelSegment(
                type=seg.type,
                start=start,
                end=end,
                start_rel=seg.start_rel,
                end_rel=seg.end_rel,
                bearing=seg.bearing,
                distance=seg.distance,
                radius=seg.radius,
                rotation=seg.rotation,
                delta=seg.delta,
                bulge=seg.bulge,
                center=center,
            ))
        return pts, export_segments

    def export_landxml(self):
        if not self.parcel_points_abs:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        path = filedialog.asksaveasfilename(parent=self, title="Save LandXML", defaultextension=".xml",
                                            filetypes=[("LandXML", "*.xml"), ("All Files", "*.*")],
                                            initialfile="parcel.xml")
        if not path:
            return
        pts, segments = self._prepare_export_geometry()
        if not pts or not segments:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        if not self._polygon_is_closed(pts):
            messagebox.showerror("Open parcel", "Parcel must close before exporting.")
            return
        try:
            root = Element("LandXML", version="1.2", xmlns="http://www.landxml.org/schema/LandXML-1.2")
            SubElement(root, "Application", name="Geo-Builder", version="1.0", desc="Deed Processor")
            SubElement(root, "Project", name="Parcel", desc=self.selected_spcs or "")
            units_elem = SubElement(root, "Units")
            unit_name = "meter" if self.settings.get("units_in", "feet").lower() == "meters" else "foot"
            SubElement(units_elem, "Linear", unit=unit_name, conversionFactor="1.0")
            cgpoints = SubElement(root, "CgPoints")
            for idx, (x, y) in enumerate(pts, start=1):
                SubElement(cgpoints, "Point", name=f"P{idx}", desc="Parcel Corner").text = f"{x:.3f} {y:.3f} 0.000"
            parcels = SubElement(root, "Parcels")
            parcel = SubElement(parcels, "Parcel", name="Parcel-1")
            coord_geom = SubElement(parcel, "CoordGeom")
            for seg in segments:
                if seg.type == "line":
                    line = SubElement(coord_geom, "Line")
                    SubElement(line, "Start").text = f"{seg.start[0]:.3f} {seg.start[1]:.3f}"
                    SubElement(line, "End").text = f"{seg.end[0]:.3f} {seg.end[1]:.3f}"
                elif seg.type == "arc":
                    rot = "cw" if (seg.rotation or "CW").upper() == "CW" else "ccw"
                    curve = SubElement(coord_geom, "Curve", rot=rot,
                                       radius=f"{seg.radius:.3f}" if seg.radius else "")
                    SubElement(curve, "Start").text = f"{seg.start[0]:.3f} {seg.start[1]:.3f}"
                    if seg.center:
                        SubElement(curve, "Center").text = f"{seg.center[0]:.3f} {seg.center[1]:.3f}"
                    SubElement(curve, "End").text = f"{seg.end[0]:.3f} {seg.end[1]:.3f}"
            tree = ElementTree(root)
            tree.write(path, encoding="utf-8", xml_declaration=True)
            self._log(f"LandXML exported → {path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_dxf(self):
        if not HAVE_EZDXF:
            messagebox.showerror("Missing ezdxf", "Install ezdxf to export DXF files.")
            return
        if not self.parcel_points_abs:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        path = filedialog.asksaveasfilename(parent=self, title="Save DXF", defaultextension=".dxf",
                                            filetypes=[("DXF", "*.dxf"), ("All Files", "*.*")],
                                            initialfile="parcel.dxf")
        if not path:
            return
        pts, segments = self._prepare_export_geometry()
        if not pts or not segments:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        if not self._polygon_is_closed(pts):
            messagebox.showerror("Open parcel", "Parcel must close before exporting.")
            return
        try:
            doc = ezdxf.new(setup=True)
            units_in = self.settings.get("units_in", "feet").lower()
            if dxf_const is not None and hasattr(doc, "header"):
                doc.header["$INSUNITS"] = dxf_const.INSUNITS_METERS if units_in == "meters" else dxf_const.INSUNITS_FEET
            elif hasattr(doc, "units") and hasattr(ezdxf, "units"):
                doc.units = ezdxf.units.M if units_in == "meters" else ezdxf.units.FOOT
            if hasattr(doc.header, "__setitem__"):
                doc.header["$PROJECTNAME"] = self.selected_spcs or "Parcel"
                if self.selected_spcs_epsg:
                    doc.header["$PROJECTDESCRIPTION"] = f"EPSG:{self.selected_spcs_epsg}"
            msp = doc.modelspace()
            vertices = []
            vertices.append((pts[0][0], pts[0][1], 0.0))
            for idx, seg in enumerate(segments):
                end = pts[idx + 1]
                bulge = seg.bulge if seg.type == "arc" else 0.0
                vertices.append((end[0], end[1], bulge))
            lwpoly = msp.add_lwpolyline(vertices, format="xyb")
            lwpoly.closed = True
            for seg in segments:
                if seg.type == "line":
                    msp.add_line(seg.start, seg.end)
                elif seg.type == "arc" and seg.center and seg.radius:
                    cx, cy = seg.center
                    start_angle = math.degrees(math.atan2(seg.start[1] - cy, seg.start[0] - cx))
                    end_angle = math.degrees(math.atan2(seg.end[1] - cy, seg.end[0] - cx))
                    if seg.rotation == "CW":
                        msp.add_arc(center=seg.center, radius=seg.radius,
                                    start_angle=end_angle, end_angle=start_angle,
                                    is_counter_clockwise=False)
                    else:
                        msp.add_arc(center=seg.center, radius=seg.radius,
                                    start_angle=start_angle, end_angle=end_angle)
            if self.selected_spcs:
                label = self.selected_spcs
                if self.selected_spcs_epsg:
                    label += f" | EPSG:{self.selected_spcs_epsg}"
                msp.add_text(label, dxfattribs={"height": 5}).set_pos((pts[0][0], pts[0][1] + 10))
            doc.saveas(path)
            self._log(f"DXF exported → {path}")
        except Exception as exc:
            messagebox.showerror("DXF export failed", str(exc))

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
                       current_tesseract_path=self.settings["tesseract_path"], current_spcs=self.selected_spcs,
                       current_ai_dir=self.settings.get("ai_training_dir", ""), on_apply=self.apply_settings)
    def apply_settings(self, mode, units_in, units_out, bearing_fmt, tess_path, ai_dir):
        previous = dict(self.settings)
        self.settings.update({"theme": mode, "units_in": units_in, "units_out": units_out, "bearing_fmt": bearing_fmt, "tesseract_path": tess_path or "", "ai_training_dir": ai_dir or ""})
        if hasattr(self, "ai_training_folder_var"):
            self.ai_training_folder_var.set(ai_dir or "")
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
        self._save_user_config()
        library_msg = ai_dir if ai_dir else "None"
        self._log(f"Settings applied → Theme={mode}, Input Units={units_in}, Output Units={units_out}, Bearing Format={bearing_fmt}, AI Library={library_msg}")
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

    def _load_user_config(self):
        if self.config_path.exists():
            try:
                self._config_parser.read(self.config_path)
            except Exception:
                self._config_parser = configparser.ConfigParser()
        section = self._config_parser.setdefault("GeoBuilder", {})
        self._user_config = dict(section)

    def _save_user_config(self):
        section = self._config_parser.setdefault("GeoBuilder", {})
        section.update({
            "units_in": self.settings.get("units_in", "feet"),
            "units_out": self.settings.get("units_out", "feet"),
            "bearing_fmt": self.settings.get("bearing_fmt", "dms"),
            "tesseract_path": self.settings.get("tesseract_path", ""),
            "spcs_name": self.selected_spcs or "",
            "spcs_epsg": str(self.selected_spcs_epsg or ""),
            "origin_easting": self.origin_easting_var.get(),
            "origin_northing": self.origin_northing_var.get(),
            "source_epsg": self.source_epsg_var.get(),
            "apply_pyproj": str(bool(self.apply_pyproj_var.get())).lower(),
            "ai_training_dir": self.ai_training_folder_var.get() if hasattr(self, "ai_training_folder_var") else "",
        })
        try:
            with self.config_path.open("w", encoding="utf-8") as fh:
                self._config_parser.write(fh)
        except Exception:
            pass

    @staticmethod
    def _coerce_int(value):
        if value in (None, ""):
            return None
        try:
            return int(str(value))
        except ValueError:
            return None

    @staticmethod
    def _to_float(value):
        if value in (None, ""):
            return None
        try:
            return float(str(value).replace(",", ""))
        except ValueError:
            return None

    def open_spcs_dialog(self):
        dialog = tk.Toplevel(self)
        dialog.title("Select State Plane Coordinate System")
        dialog.transient(self)
        dialog.grab_set()
        tk.Label(dialog, text="Search:").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        search_var = tk.StringVar()
        search_entry = tk.Entry(dialog, textvariable=search_var, width=40)
        search_entry.grid(row=0, column=1, padx=6, pady=6, sticky="we")
        options = sorted(SPCS_ZONES.keys())
        listbox = tk.Listbox(dialog, height=12, exportselection=False)
        listbox.grid(row=1, column=0, columnspan=2, padx=6, pady=6, sticky="nsew")
        for opt in options:
            listbox.insert(tk.END, opt)
        if self.selected_spcs in options:
            idx = options.index(self.selected_spcs)
            listbox.selection_set(idx)
            listbox.see(idx)
        dialog.columnconfigure(1, weight=1)
        dialog.rowconfigure(1, weight=1)

        def apply_filter(*_):
            term = search_var.get().lower()
            listbox.delete(0, tk.END)
            matches = [opt for opt in options if term in opt.lower()]
            for opt in (matches or options):
                listbox.insert(tk.END, opt)
            if matches:
                listbox.selection_clear(0, tk.END)
                listbox.selection_set(0)

        search_var.trace_add("write", apply_filter)

        def choose():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("No selection", "Choose a coordinate system.", parent=dialog)
                return
            value = listbox.get(sel[0])
            self.selected_spcs = value
            self.selected_spcs_epsg = SPCS_ZONES.get(value)
            self.settings["spcs_name"] = value
            self.settings["spcs_epsg"] = self.selected_spcs_epsg
            if getattr(self, "spcs_value_lbl", None):
                self.spcs_value_lbl.config(text=value)
            self._save_user_config()
            self._log(f"Selected SPCS: {value}")
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=6)
        tk.Button(btn_frame, text="OK", command=choose).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=4)
        search_entry.focus_set()
        dialog.wait_window(dialog)

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
