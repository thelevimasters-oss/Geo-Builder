"""Deed processor GUI with LandXML and DXF export support."""

from __future__ import annotations

import configparser
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:  # Optional DXF support
    import ezdxf
    from ezdxf import units as dxf_units

    HAVE_EZDXF = True
except Exception:  # pragma: no cover - handled gracefully at runtime
    ezdxf = None  # type: ignore
    dxf_units = None  # type: ignore
    HAVE_EZDXF = False

try:  # Optional coordinate transformations
    import pyproj

    HAVE_PYPROJ = True
except Exception:  # pragma: no cover
    pyproj = None  # type: ignore
    HAVE_PYPROJ = False

try:  # Optional parcel preview
    import matplotlib
    matplotlib.use("Agg")  # Use a non-interactive backend for compatibility
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure

    HAVE_MPL = True
except Exception:  # pragma: no cover
    Figure = None  # type: ignore
    FigureCanvasTkAgg = None  # type: ignore
    HAVE_MPL = False

# --------------------------------------------------------------------------------------------------
# Data structures
# --------------------------------------------------------------------------------------------------


@dataclass
class ParsedCall:
    """Internal representation of a deed call."""

    type: str
    bearing: Optional[str] = None
    distance: Optional[float] = None
    radius: Optional[float] = None
    arc_length: Optional[float] = None
    chord_length: Optional[float] = None
    chord_bearing: Optional[str] = None
    rotation: Optional[str] = None  # "LEFT"/"RIGHT" or "CCW"/"CW"


@dataclass
class ParcelSegment:
    """Computed geometric segment (line or arc)."""

    type: str
    start: Tuple[float, float]
    end: Tuple[float, float]
    start_rel: Tuple[float, float]
    end_rel: Tuple[float, float]
    bearing: Optional[str] = None
    distance: Optional[float] = None
    radius: Optional[float] = None
    rotation: Optional[str] = None
    delta: Optional[float] = None  # degrees
    bulge: float = 0.0
    center: Optional[Tuple[float, float]] = None


# --------------------------------------------------------------------------------------------------
# Library of State Plane coordinate systems (NAD83) — at least one per state (>=50 entries)
# --------------------------------------------------------------------------------------------------

SPCS_ZONES: Dict[str, int] = {
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


class ToolTip:
    """Lightweight tooltip implementation for Tk widgets."""

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
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#ffffe0",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Segoe UI", 9),
        )
        label.pack(ipadx=4, ipady=2)

    def hide(self, _event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


@dataclass
class ProcessResult:
    points_rel: List[Tuple[float, float]]
    points_abs: List[Tuple[float, float]]
    segments: List[ParcelSegment]


class DeedProcessorApp(tk.Tk):
    """Tk GUI that parses deed calls, computes parcel points, and exports to LandXML/DXF."""

    CONFIG_SECTION = "DeedProcessor"

    def __init__(self) -> None:
        super().__init__()
        self.title("Geo-Builder Deed Processor")
        self.geometry("1100x720")

        self.config_path = Path.home() / ".geo_builder.ini"
        self.config_parser = configparser.ConfigParser()
        self._load_config()

        self.calls: List[ParsedCall] = []
        self.segments: List[ParcelSegment] = []
        self.relative_points: List[Tuple[float, float]] = []
        self.offset_points: List[Tuple[float, float]] = []

        self.units_var = tk.StringVar(
            value=self.config_parser.get(self.CONFIG_SECTION, "units", fallback="feet")
        )
        self.origin_easting_var = tk.StringVar(value=self.config_parser.get(self.CONFIG_SECTION, "origin_easting", fallback="0.0"))
        self.origin_northing_var = tk.StringVar(value=self.config_parser.get(self.CONFIG_SECTION, "origin_northing", fallback="0.0"))
        self.source_epsg_var = tk.StringVar(value=self.config_parser.get(self.CONFIG_SECTION, "source_epsg", fallback=""))
        self.apply_pyproj_var = tk.BooleanVar(value=self.config_parser.getboolean(self.CONFIG_SECTION, "apply_pyproj", fallback=False))

        spcs_name = self.config_parser.get(self.CONFIG_SECTION, "spcs_name", fallback="")
        self.selected_spcs = tk.StringVar(value=spcs_name)
        self.selected_spcs_epsg: Optional[int] = SPCS_ZONES.get(spcs_name)

        self.status_var = tk.StringVar(value="Paste or load deed calls, then click Process Deed.")

        self._build_menus()
        self._build_layout()
        self._set_status("Ready.")

    def _build_menus(self) -> None:
        menubar = tk.Menu(self)
        settings_menu = tk.Menu(menubar, tearoff=False)
        settings_menu.add_command(label="Select Coordinate System…", command=self.open_spcs_dialog)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        self.config(menu=menubar)

    def _build_layout(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        deed_frame = ttk.LabelFrame(main, text="Deed Calls (CSV: TYPE,Bearing,Distance,Radius,ArcLength,ChordLength,ChordBearing,Rotation)")
        deed_frame.pack(side="left", fill="both", expand=True, padx=(0, 12))

        self.deed_text = tk.Text(deed_frame, width=50, height=20, wrap="word", font=("Consolas", 10))
        self.deed_text.pack(fill="both", expand=True)
        ToolTip(
            self.deed_text,
            "Enter one call per line. Example:\nLINE,N45E,100\nCURVE,RIGHT,,120,,85,N30E,RIGHT",
        )

        control_frame = ttk.Frame(main)
        control_frame.pack(side="right", fill="both", expand=True)

        origin_frame = ttk.LabelFrame(control_frame, text="Point of Beginning / Origin")
        origin_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(origin_frame, text="Easting / X:").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        e_entry = ttk.Entry(origin_frame, textvariable=self.origin_easting_var, width=18)
        e_entry.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        ToolTip(e_entry, "Absolute easting/X for the parcel's point of beginning.")

        ttk.Label(origin_frame, text="Northing / Y:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        n_entry = ttk.Entry(origin_frame, textvariable=self.origin_northing_var, width=18)
        n_entry.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        ToolTip(n_entry, "Absolute northing/Y for the parcel's point of beginning.")

        units_frame = ttk.Frame(origin_frame)
        units_frame.grid(row=0, column=2, rowspan=2, padx=(12, 0), sticky="nsw")
        ttk.Label(units_frame, text="Units:").pack(anchor="w")
        units_menu = ttk.OptionMenu(units_frame, self.units_var, self.units_var.get(), "feet", "meters")
        units_menu.pack(anchor="w", pady=(0, 4))
        ToolTip(units_menu, "Working linear units for distances.")

        proj_frame = ttk.LabelFrame(control_frame, text="Coordinate Reference System")
        proj_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(proj_frame, text="Selected SPCS:").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.spcs_label = ttk.Label(proj_frame, text=self.selected_spcs.get() or "None")
        self.spcs_label.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        ttk.Button(proj_frame, text="Choose…", command=self.open_spcs_dialog).grid(row=0, column=2, padx=4, pady=2)

        ttk.Label(proj_frame, text="Source EPSG (optional):").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        src_entry = ttk.Entry(proj_frame, textvariable=self.source_epsg_var, width=12)
        src_entry.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        ToolTip(src_entry, "Provide the EPSG code of your working coordinates if you want pyproj to transform to the SPCS zone.")

        transform_chk = ttk.Checkbutton(
            proj_frame,
            text="Transform with pyproj when exporting",
            variable=self.apply_pyproj_var,
            state="normal" if HAVE_PYPROJ else "disabled",
        )
        transform_chk.grid(row=2, column=0, columnspan=3, sticky="w", padx=4, pady=2)
        if not HAVE_PYPROJ:
            ToolTip(transform_chk, "Install pyproj to enable automatic transformations.")

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill="x", pady=(4, 10))

        process_btn = ttk.Button(button_frame, text="Process Deed", command=self.process_deed)
        process_btn.pack(side="left", padx=4)
        ToolTip(process_btn, "Parse calls, compute relative/absolute parcel coordinates, and update the preview.")

        export_xml_btn = ttk.Button(button_frame, text="Export LandXML", command=self.export_landxml)
        export_xml_btn.pack(side="left", padx=4)
        ToolTip(export_xml_btn, "Write the parcel geometry to a LandXML file using the selected origin.")

        export_dxf_btn = ttk.Button(button_frame, text="Export DXF", command=self.export_dxf)
        export_dxf_btn.pack(side="left", padx=4)
        ToolTip(export_dxf_btn, "Create a DXF with parcel lines/arcs. Requires ezdxf.")

        points_frame = ttk.LabelFrame(control_frame, text="Computed Parcel Points")
        points_frame.pack(fill="both", expand=True)

        columns = ("index", "rel_x", "rel_y", "abs_x", "abs_y")
        self.points_tree = ttk.Treeview(points_frame, columns=columns, show="headings", height=8)
        for col, label in zip(columns, ["#", "ΔE", "ΔN", "Easting", "Northing"]):
            self.points_tree.heading(col, text=label)
            anchor = "center" if col == "index" else "e"
            self.points_tree.column(col, width=90, anchor=anchor)
        self.points_tree.pack(fill="both", expand=True)

        preview_frame = ttk.LabelFrame(control_frame, text="Parcel Preview")
        preview_frame.pack(fill="both", expand=True, pady=(8, 0))

        if HAVE_MPL:
            self.figure = Figure(figsize=(4, 3), dpi=100)
            self.ax = self.figure.add_subplot(111)
            self.ax.set_title("Plan View")
            self.ax.set_aspect("equal", adjustable="datalim")
            self.ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)
            self.canvas = FigureCanvasTkAgg(self.figure, master=preview_frame)
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            ttk.Label(preview_frame, text="Install matplotlib for live previews.", foreground="#777").pack(fill="both", expand=True, padx=8, pady=8)

        status_bar = ttk.Label(self, textvariable=self.status_var, anchor="w", padding=8)
        status_bar.pack(fill="x", side="bottom")

    def _load_config(self) -> None:
        if self.config_path.exists():
            try:
                self.config_parser.read(self.config_path)
            except Exception:
                self.config_parser = configparser.ConfigParser()
        if self.CONFIG_SECTION not in self.config_parser:
            self.config_parser[self.CONFIG_SECTION] = {}

    def _save_config(self) -> None:
        section = self.config_parser.setdefault(self.CONFIG_SECTION, {})
        section["spcs_name"] = self.selected_spcs.get()
        section["units"] = self.units_var.get()
        section["origin_easting"] = self.origin_easting_var.get()
        section["origin_northing"] = self.origin_northing_var.get()
        section["source_epsg"] = self.source_epsg_var.get()
        section["apply_pyproj"] = str(self.apply_pyproj_var.get())
        with self.config_path.open("w", encoding="utf-8") as f:
            self.config_parser.write(f)

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)

    def _get_origin(self) -> Tuple[float, float]:
        try:
            easting = float(self.origin_easting_var.get() or 0.0)
            northing = float(self.origin_northing_var.get() or 0.0)
        except ValueError:
            raise ValueError("Origin coordinates must be numeric values.")
        return easting, northing

    def process_deed(self) -> None:
        try:
            origin = self._get_origin()
        except ValueError as exc:
            messagebox.showerror("Invalid origin", str(exc))
            return

        text = self.deed_text.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("No deed calls", "Enter deed calls before processing.")
            return

        try:
            self.calls = self._parse_calls(text)
        except ValueError as exc:
            messagebox.showerror("Parsing error", str(exc))
            return

        if not self.calls:
            messagebox.showwarning("Empty deed", "No valid calls were found.")
            return

        result = self._compute_geometry(self.calls, origin)
        self.relative_points = result.points_rel
        self.offset_points = result.points_abs
        self.segments = result.segments

        closed = self._ensure_closed()
        if not closed:
            messagebox.showwarning("Open parcel", "The parcel did not close; a closing segment was added.")

        self._update_points_table()
        self._update_preview()
        self._save_config()
        self._set_status(f"Processed {len(self.calls)} calls. Parcel {'closed' if closed else 'forced closed'}.")

    def _ensure_closed(self) -> bool:
        if not self.offset_points:
            return True
        start = self.offset_points[0]
        end = self.offset_points[-1]
        distance = math.hypot(end[0] - start[0], end[1] - start[1])
        if distance <= 1e-4:
            return True
        self.relative_points.append(self.relative_points[0])
        self.offset_points.append(start)
        seg = ParcelSegment(
            type="line",
            start=self.offset_points[-2],
            end=start,
            start_rel=self.relative_points[-2],
            end_rel=self.relative_points[0],
            bearing=None,
            distance=distance,
        )
        self.segments.append(seg)
        return False

    def _polygon_is_closed(self, pts: Sequence[Tuple[float, float]], tol: float = 1e-4) -> bool:
        if len(pts) < 3:
            return False
        start = pts[0]
        end = pts[-1]
        return math.hypot(end[0] - start[0], end[1] - start[1]) <= tol

    def _parse_calls(self, text: str) -> List[ParsedCall]:
        calls: List[ParsedCall] = []
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            # Accept JSON dictionaries as well as CSV style rows
            if line.startswith("{") and line.endswith("}"):
                data = json.loads(line)
                call = ParsedCall(
                    type=str(data.get("type") or data.get("Type") or "").strip(),
                    bearing=data.get("bearing") or data.get("Bearing"),
                    distance=self._to_float(data.get("distance") or data.get("Distance")),
                    radius=self._to_float(data.get("radius") or data.get("Radius")),
                    arc_length=self._to_float(data.get("arc_length") or data.get("Arc Length")),
                    chord_length=self._to_float(data.get("chord_length") or data.get("Chord Length")),
                    chord_bearing=data.get("chord_bearing") or data.get("Chord Bearing"),
                    rotation=(data.get("rotation") or data.get("Rotation") or "").upper(),
                )
            else:
                parts = [p.strip() for p in line.split(",")]
                if len(parts) < 1:
                    continue
                ctype = parts[0].upper()
                if ctype not in {"LINE", "CURVE"}:
                    raise ValueError(f"Unsupported call type '{parts[0]}' in line: {line}")
                bearing = parts[1] if len(parts) > 1 and parts[1] else None
                distance = self._to_float(parts[2]) if len(parts) > 2 else None
                radius = self._to_float(parts[3]) if len(parts) > 3 else None
                arc_length = self._to_float(parts[4]) if len(parts) > 4 else None
                chord_length = self._to_float(parts[5]) if len(parts) > 5 else None
                chord_bearing = parts[6] if len(parts) > 6 and parts[6] else None
                rotation = (parts[7] if len(parts) > 7 else None) or None
                call = ParsedCall(
                    type=ctype,
                    bearing=bearing,
                    distance=distance,
                    radius=radius,
                    arc_length=arc_length,
                    chord_length=chord_length,
                    chord_bearing=chord_bearing,
                    rotation=(rotation or "").upper() if rotation else None,
                )
            if not call.type:
                continue
            calls.append(call)
        return calls

    def _compute_geometry(self, calls: Sequence[ParsedCall], origin: Tuple[float, float]) -> ProcessResult:
        points_rel: List[Tuple[float, float]] = [(0.0, 0.0)]
        points_abs: List[Tuple[float, float]] = [origin]
        segments: List[ParcelSegment] = []

        current_dir: Optional[float] = None  # radians, east=0 ccw
        for call in calls:
            start_rel = points_rel[-1]
            start_abs = points_abs[-1]
            if call.type.upper() == "LINE":
                if call.distance is None or (call.distance or 0) == 0:
                    raise ValueError("Line calls require a numeric distance.")
                if not call.bearing:
                    raise ValueError("Line calls require a bearing (quadrant or azimuth).")
                theta = self._bearing_to_radians(call.bearing)
                current_dir = theta
                dx = math.cos(theta) * call.distance
                dy = math.sin(theta) * call.distance
                end_rel = (start_rel[0] + dx, start_rel[1] + dy)
                end_abs = (start_abs[0] + dx, start_abs[1] + dy)
                segment = ParcelSegment(
                    type="line",
                    start=start_abs,
                    end=end_abs,
                    start_rel=start_rel,
                    end_rel=end_rel,
                    bearing=call.bearing,
                    distance=call.distance,
                )
                segments.append(segment)
                points_rel.append(end_rel)
                points_abs.append(end_abs)
            elif call.type.upper() == "CURVE":
                if call.radius is None and call.chord_length is None:
                    raise ValueError("Curve calls require a radius or chord length.")
                rotation = (call.rotation or "LEFT").upper()
                if rotation in {"LEFT", "CCW"}:
                    rotation = "CCW"
                elif rotation in {"RIGHT", "CW"}:
                    rotation = "CW"
                else:
                    rotation = "CCW"

                radius = call.radius or (
                    (call.chord_length or 0.0) / (2 * math.sin(math.radians(45)))
                )
                chord = call.chord_length
                if chord is None and call.arc_length and radius:
                    chord = 2 * radius * math.sin((call.arc_length / radius) / 2.0)
                if chord is None:
                    raise ValueError("Unable to determine chord length for curve.")
                if chord / 2.0 > radius:
                    raise ValueError("Invalid curve: chord exceeds diameter for given radius.")

                delta = call.arc_length / radius if (call.arc_length and radius) else 2 * math.asin(chord / (2 * radius))
                delta_deg = math.degrees(delta)

                if call.chord_bearing:
                    chord_angle = self._bearing_to_radians(call.chord_bearing)
                elif current_dir is not None:
                    turn = delta if rotation == "CCW" else -delta
                    chord_angle = current_dir + turn / 2.0
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

                segment = ParcelSegment(
                    type="arc",
                    start=start_abs,
                    end=end_abs,
                    start_rel=start_rel,
                    end_rel=end_rel,
                    radius=radius,
                    rotation=rotation,
                    delta=delta_deg,
                    bulge=bulge,
                    center=center,
                )
                segments.append(segment)
                points_rel.append(end_rel)
                points_abs.append(end_abs)
            else:
                raise ValueError(f"Unsupported call type: {call.type}")

        return ProcessResult(points_rel, points_abs, segments)

    def _compute_arc_center(
        self,
        start: Tuple[float, float],
        end: Tuple[float, float],
        radius: float,
        rotation: str,
    ) -> Optional[Tuple[float, float]]:
        dx = end[0] - start[0]
        dy = end[1] - start[1]
        chord = math.hypot(dx, dy)
        if chord == 0 or radius <= 0:
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

    def _to_float(self, value: Optional[str]) -> Optional[float]:
        if value in (None, ""):
            return None
        try:
            return float(str(value))
        except ValueError:
            return None

    def _bearing_to_radians(self, bearing: str) -> float:
        s = bearing.strip().upper()
        s = s.replace("DEG", "°").replace("º", "°").replace("\\", "").replace("D", "°")
        quad_match = re.match(r"^([NS])\s*([0-9.]+)(?:°\s*([0-9.]+))?(?:'\s*([0-9.]+))?\s*([EW])$", s)
        if quad_match:
            ns, deg, minutes, seconds, ew = quad_match.groups()
            deg = float(deg)
            minutes = float(minutes) if minutes else 0.0
            seconds = float(seconds) if seconds else 0.0
            angle = deg + minutes / 60.0 + seconds / 3600.0
            if ns == "N" and ew == "E":
                az = angle
            elif ns == "N" and ew == "W":
                az = 360.0 - angle
            elif ns == "S" and ew == "E":
                az = 180.0 - angle
            else:
                az = 180.0 + angle
            return math.radians((90.0 - az) % 360.0)
        try:
            az = float(re.sub(r"[^0-9.+-]", "", s))
        except ValueError:
            raise ValueError(f"Cannot parse bearing '{bearing}'.")
        return math.radians((90.0 - az) % 360.0)

    def _update_points_table(self) -> None:
        for child in self.points_tree.get_children():
            self.points_tree.delete(child)
        for idx, (rel, abs_pt) in enumerate(zip(self.relative_points, self.offset_points), start=1):
            self.points_tree.insert(
                "",
                "end",
                values=(
                    idx,
                    f"{rel[0]:.3f}",
                    f"{rel[1]:.3f}",
                    f"{abs_pt[0]:.3f}",
                    f"{abs_pt[1]:.3f}",
                ),
            )

    def _update_preview(self) -> None:
        if not HAVE_MPL:
            return
        self.ax.clear()
        self.ax.set_title("Plan View")
        self.ax.set_aspect("equal", adjustable="datalim")
        self.ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)
        if len(self.offset_points) >= 2:
            xs = [p[0] for p in self.offset_points]
            ys = [p[1] for p in self.offset_points]
            self.ax.plot(xs, ys, marker="o")
        self.canvas.draw_idle()

    def open_spcs_dialog(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Select State Plane Coordinate System")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="Search:").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(dialog, textvariable=search_var, width=40)
        search_entry.grid(row=0, column=1, padx=6, pady=6, sticky="we")

        options = sorted(SPCS_ZONES.keys())
        filtered_options = tk.StringVar(value=options)

        listbox = tk.Listbox(dialog, listvariable=filtered_options, height=12, exportselection=False)
        listbox.grid(row=1, column=0, columnspan=2, padx=6, pady=6, sticky="nsew")

        if self.selected_spcs.get() in options:
            idx = options.index(self.selected_spcs.get())
            listbox.selection_set(idx)
            listbox.see(idx)

        dialog.columnconfigure(1, weight=1)
        dialog.rowconfigure(1, weight=1)

        def apply_filter(*_args):
            term = search_var.get().lower()
            matches = [opt for opt in options if term in opt.lower()]
            filtered_options.set(matches or options)
            listbox.delete(0, tk.END)
            for opt in matches or options:
                listbox.insert(tk.END, opt)
            if matches:
                listbox.selection_clear(0, tk.END)
                listbox.selection_set(0)

        search_var.trace_add("write", apply_filter)

        def choose():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No selection", "Choose a coordinate system.", parent=dialog)
                return
            value = listbox.get(selection[0])
            self.selected_spcs.set(value)
            self.selected_spcs_epsg = SPCS_ZONES.get(value)
            self.spcs_label.config(text=value)
            dialog.destroy()
            self._save_config()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=6)
        ttk.Button(btn_frame, text="OK", command=choose).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=4)

        search_entry.focus_set()
        dialog.wait_window(dialog)

    def _prepare_export_geometry(self) -> Tuple[List[Tuple[float, float]], List[ParcelSegment]]:
        if not self.offset_points:
            return [], []

        pts = list(self.offset_points)
        transformer = None
        if self.apply_pyproj_var.get() and HAVE_PYPROJ and self.selected_spcs_epsg and self.source_epsg_var.get():
            try:
                src_epsg = int(self.source_epsg_var.get())
                dst_epsg = self.selected_spcs_epsg
                transformer = pyproj.Transformer.from_crs(
                    f"EPSG:{src_epsg}", f"EPSG:{dst_epsg}", always_xy=True
                )
                pts = [transformer.transform(x, y) for x, y in pts]
            except Exception as exc:
                transformer = None
                messagebox.showwarning("Transformation failed", str(exc))

        export_segments: List[ParcelSegment] = []
        for idx, seg in enumerate(self.segments):
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
            export_segments.append(
                ParcelSegment(
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
                )
            )

        return pts, export_segments

    def export_landxml(self) -> None:
        if not self.offset_points:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        path = filedialog.asksaveasfilename(
            title="Save LandXML",
            defaultextension=".xml",
            filetypes=[("LandXML", "*.xml"), ("All Files", "*.*")],
            initialfile="parcel.xml",
        )
        if not path:
            return
        pts, segments = self._prepare_export_geometry()
        if not pts or not segments:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        if not self._polygon_is_closed(pts):
            messagebox.showerror("Open parcel", "The parcel geometry must close before exporting.")
            return
        try:
            import xml.etree.ElementTree as ET

            root = ET.Element("LandXML", version="1.2", xmlns="http://www.landxml.org/schema/LandXML-1.2")
            ET.SubElement(root, "Application", name="Geo-Builder", version="1.0", desc="Deed Processor")
            project = ET.SubElement(root, "Project", name="Parcel", desc=self.selected_spcs.get() or "")
            units = ET.SubElement(root, "Units")
            unit_name = {"feet": "foot", "meters": "meter"}.get(self.units_var.get().lower(), "foot")
            ET.SubElement(units, "Linear", unit=unit_name, conversionFactor="1.0")
            cgpoints = ET.SubElement(root, "CgPoints")
            for idx, (x, y) in enumerate(pts, start=1):
                ET.SubElement(cgpoints, "Point", name=f"P{idx}", desc="Parcel Corner").text = f"{x:.3f} {y:.3f} 0.000"
            parcels = ET.SubElement(root, "Parcels")
            parcel = ET.SubElement(parcels, "Parcel", name="Parcel-1")
            if pts:
                begin_x, begin_y = pts[0]
                parcel.set("begin.x", f"{begin_x:.3f}")
                parcel.set("begin.y", f"{begin_y:.3f}")
            coord_geom = ET.SubElement(parcel, "CoordGeom")
            for seg in segments:
                if seg.type == "line":
                    line = ET.SubElement(coord_geom, "Line")
                    ET.SubElement(line, "Start").text = f"{seg.start[0]:.3f} {seg.start[1]:.3f}"
                    ET.SubElement(line, "End").text = f"{seg.end[0]:.3f} {seg.end[1]:.3f}"
                elif seg.type == "arc":
                    rot = "cw" if (seg.rotation or "CW").upper() == "CW" else "ccw"
                    curve = ET.SubElement(coord_geom, "Curve", rot=rot, radius=f"{seg.radius:.3f}" if seg.radius else "")
                    ET.SubElement(curve, "Start").text = f"{seg.start[0]:.3f} {seg.start[1]:.3f}"
                    if seg.center:
                        ET.SubElement(curve, "Center").text = f"{seg.center[0]:.3f} {seg.center[1]:.3f}"
                    ET.SubElement(curve, "End").text = f"{seg.end[0]:.3f} {seg.end[1]:.3f}"
            tree = ET.ElementTree(root)
            tree.write(path, encoding="utf-8", xml_declaration=True)
            self._set_status(f"LandXML exported → {path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_dxf(self) -> None:
        if not HAVE_EZDXF:
            messagebox.showerror("Missing ezdxf", "Install ezdxf to export DXF files.")
            return
        if not self.offset_points:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        path = filedialog.asksaveasfilename(
            title="Save DXF",
            defaultextension=".dxf",
            filetypes=[("DXF", "*.dxf"), ("All Files", "*.*")],
            initialfile="parcel.dxf",
        )
        if not path:
            return
        pts, segments = self._prepare_export_geometry()
        if not pts or not segments:
            messagebox.showerror("No geometry", "Process the deed before exporting.")
            return
        if not self._polygon_is_closed(pts):
            messagebox.showerror("Open parcel", "The parcel geometry must close before exporting.")
            return
        try:
            doc = ezdxf.new(setup=True)
            if self.units_var.get() == "meters":
                doc.units = dxf_units.M
            else:
                doc.units = dxf_units.FOOT
            project_label = self.selected_spcs.get() or "Parcel"
            doc.header["$PROJECTNAME"] = project_label
            if self.selected_spcs_epsg:
                doc.header["$PROJECTDESCRIPTION"] = f"EPSG:{self.selected_spcs_epsg}"
            msp = doc.modelspace()

            # Build LWPolyline with bulge values for arcs
            if not pts:
                raise ValueError("No points to export")
            vertices: List[Tuple[float, float, float]] = []
            vertices.append((pts[0][0], pts[0][1], 0.0))
            for idx, seg in enumerate(segments):
                end = pts[idx + 1]
                bulge = seg.bulge if seg.type == "arc" else 0.0
                vertices.append((end[0], end[1], bulge))
            lwpoly = msp.add_lwpolyline(vertices, format="xyb")
            lwpoly.closed = True

            # Also add explicit line and arc entities for compatibility with some consumers
            for seg in segments:
                if seg.type == "line":
                    msp.add_line(seg.start, seg.end)
                elif seg.type == "arc" and seg.center and seg.radius:
                    cx, cy = seg.center
                    start_angle = math.degrees(math.atan2(seg.start[1] - cy, seg.start[0] - cx))
                    end_angle = math.degrees(math.atan2(seg.end[1] - cy, seg.end[0] - cx))
                    if seg.rotation == "CW":
                        msp.add_arc(
                            center=seg.center,
                            radius=seg.radius,
                            start_angle=end_angle,
                            end_angle=start_angle,
                            is_counter_clockwise=False,
                        )
                    else:
                        msp.add_arc(
                            center=seg.center,
                            radius=seg.radius,
                            start_angle=start_angle,
                            end_angle=end_angle,
                        )

            # Metadata text entity
            info = self.selected_spcs.get()
            if info:
                label = info
                if self.selected_spcs_epsg:
                    label = f"{info} | EPSG:{self.selected_spcs_epsg}"
                msp.add_text(label, dxfattribs={"height": 5}).set_pos((pts[0][0], pts[0][1] + 10))

            doc.saveas(path)
            self._set_status(f"DXF exported → {path}")
        except Exception as exc:
            messagebox.showerror("DXF export failed", str(exc))

if __name__ == "__main__":
    app = DeedProcessorApp()
    app.mainloop()
