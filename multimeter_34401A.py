"""
HP/Agilent 34401A Multimeter – GUI
===================================
Features:
  - Device configuration (function, range, resolution)
  - Active function and range visible in display
  - Continuous measurement with configurable interval
  - Real-time plot with optional autoscaling
  - Excel export (.xlsx) with live per-value saving
  - Settings persisted in settings.json (auto-connect on startup)
  - Language: English / Deutsch  |  Theme: Dark / Bright
  - Gear-icon settings dialog

Requirements:
  pip install pyvisa matplotlib openpyxl numpy
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import datetime
import queue
import random
import json
import os
import csv

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np

try:
    import pyvisa
    PYVISA_AVAILABLE = True
except ImportError:
    PYVISA_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# Localisation strings
# ─────────────────────────────────────────────────────────────────────────────

STRINGS: dict[str, dict[str, str]] = {
    "de": {
        "title":          "HP/Agilent 34401A Multimeter – Messsystem",
        "connection":     "🔌 Verbindung",
        "interface":      "Schnittstelle:",
        "connect":        "Verbinden",
        "disconnect":     "Trennen",
        "disconnected":   "● Getrennt",
        "connected":      "● Verbunden",
        "simulation":     "● Simulation",
        "functions":      "⚙ Gerätefunktionen",
        "meas_func":      "Messfunktion:",
        "meas_range":     "Messbereich:",
        "resolution":     "Auflösung:",
        "acquisition":    "⏱ Aufnahme",
        "interval":       "Intervall (ms):",
        "max_points":     "Max. Punkte:",
        "autoscale":      "Autoskalierung",
        "scroll":         "Auto-Scroll",
        "scroll_win":     "Fenster (s):",
        "show_stats":     "Statistik anzeigen",
        "file":           "💾 Datei",
        "save_dir":       "Speicherpfad:",
        "prefix_lbl":     "Prefix (ersetzt 'Messung'):",
        "suffix_lbl":     "Suffix:",
        "suffix_dt":      "Datum / Uhrzeit",
        "suffix_cnt":     "Lfd. Nummer",
        "preview":        "Vorschau:",
        "autosave_lbl":   "🔄 Jeden Messwert sofort speichern",
        "save_now":       "💾 Jetzt in Excel speichern",
        "clear_data":     "🗑 Daten löschen",
        "save_settings":  "⚙ Einstellungen speichern",
        "start":          "▶  Start",
        "stop":           "⏹  Stop",
        "ready":          "Bereit – Simulationsmodus",
        "range_lbl":      "Bereich:",
        "meas_running":   "Messung läuft …",
        "meas_stopped":   "Messung gestoppt",
        "pts_recorded":   "Punkte aufgenommen",
        "max_pts_reached":"Max. Punkte erreicht",
        "no_data":        "Keine Messdaten vorhanden.",
        "saved":          "Datei gespeichert:\n",
        "save_error":     "Speicherfehler",
        "data_cleared":   "Daten gelöscht",
        "connecting":     "Verbinde automatisch mit",
        "auto_connected": "Automatisch verbunden:",
        "auto_conn_fail": "Auto-Verbindung fehlgeschlagen – Simulationsmodus aktiv",
        "live_active":    "🔄 Live-Speicherung aktiv  →  ",
        "live_done":      "✔ Gespeichert",
        "live_err":       "⚠ Fehler: ",
        "settings_title": "Einstellungen",
        "lang_lbl":       "Sprache:",
        "theme_lbl":      "Design:",
        "theme_dark":     "Dunkel",
        "theme_bright":   "Hell",
        "apply":          "Anwenden & Neustart",
        "devices_found":  "Gerät(e) gefunden",
        "no_devices":     "Keine Geräte – Simulationsmodus verfügbar",
        "settings_saved": "✔ Einstellungen gespeichert  →  ",
        "stop_first":     "Bitte zuerst Messung stoppen.",
        "min_interval":   "Mindestintervall: 50 ms",
        "input_error":    "Eingabefehler",
        "openpyxl_missing": "Modul 'openpyxl' fehlt.\npip install openpyxl",
        "pt_count":       "Punkte",
        "statistics":     "Statistik",
        "mean":           "Mittelwert (μ):",
        "std":            "Std.-Abw. (σ):",
        "minimum":        "Minimum:",
        "maximum":        "Maximum:",
        "peak_peak":      "Peak-Peak:",
        "count":          "Anzahl:",
        "date_time":      "Datum / Zeit:",
        "meas_function":  "Messfunktion:",
        "meas_range_hdr": "Messbereich:",
        "resolution_hdr": "Auflösung:",
        "interval_hdr":   "Intervall (ms):",
        "num_points":     "Anzahl Punkte:",
        "col_nr":         "#",
        "col_time":       "Zeit (s)",
        "col_abs":        "Uhrzeit",
        "chart_title":    "Messverlauf",
        "time_axis":      "Zeit (s)",
        "warn_stop_first": "Warnung",
        "file_missing":   "Dateiname fehlt",
        "enter_filename":  "Bitte Dateinamen eingeben.",
        "finalize_err":   "⚠ Finalisierung fehlgeschlagen: ",
        "file_open_err":  "⚠ Datei konnte nicht angelegt werden: ",
        "auto_save_info":   "🔄 Autosave beim Messende aktiv",
        "excel_stats":  "Statistik in Excel",
        "excel_chart":  "Diagramm in Excel",
        "sheet_data":   "Messdaten",
        "sheet_chart":  "Diagramm",
    },
    "en": {
        "title":          "HP/Agilent 34401A Multimeter – Measurement System",
        "connection":     "🔌 Connection",
        "interface":      "Interface:",
        "connect":        "Connect",
        "disconnect":     "Disconnect",
        "disconnected":   "● Disconnected",
        "connected":      "● Connected",
        "simulation":     "● Simulation",
        "functions":      "⚙ Device Functions",
        "meas_func":      "Function:",
        "meas_range":     "Range:",
        "resolution":     "Resolution:",
        "acquisition":    "⏱ Acquisition",
        "interval":       "Interval (ms):",
        "max_points":     "Max. Points:",
        "autoscale":      "Autoscale",
        "scroll":         "Auto-Scroll",
        "scroll_win":     "Window (s):",
        "show_stats":     "Show statistics",
        "file":           "💾 File",
        "save_dir":       "Save directory:",
        "prefix_lbl":     "Prefix (replaces 'Measurement'):",
        "suffix_lbl":     "Suffix:",
        "suffix_dt":      "Date / Time",
        "suffix_cnt":     "Sequential No.",
        "preview":        "Preview:",
        "autosave_lbl":   "🔄 Save every measurement immediately",
        "save_now":       "💾 Save to Excel now",
        "clear_data":     "🗑 Clear data",
        "save_settings":  "⚙ Save settings",
        "start":          "▶  Start",
        "stop":           "⏹  Stop",
        "ready":          "Ready – Simulation mode",
        "range_lbl":      "Range:",
        "meas_running":   "Measurement running …",
        "meas_stopped":   "Measurement stopped",
        "pts_recorded":   "points recorded",
        "max_pts_reached":"Max. points reached",
        "no_data":        "No measurement data available.",
        "saved":          "File saved:\n",
        "save_error":     "Save error",
        "data_cleared":   "Data cleared",
        "connecting":     "Auto-connecting to",
        "auto_connected": "Auto-connected:",
        "auto_conn_fail": "Auto-connect failed – simulation mode active",
        "live_active":    "🔄 Live saving active  →  ",
        "live_done":      "✔ Saved",
        "live_err":       "⚠ Error: ",
        "settings_title": "Settings",
        "lang_lbl":       "Language:",
        "theme_lbl":      "Theme:",
        "theme_dark":     "Dark",
        "theme_bright":   "Bright",
        "apply":          "Apply & Restart",
        "devices_found":  "device(s) found",
        "no_devices":     "No devices – simulation mode available",
        "settings_saved": "✔ Settings saved  →  ",
        "stop_first":     "Please stop measurement first.",
        "min_interval":   "Minimum interval: 50 ms",
        "input_error":    "Input error",
        "openpyxl_missing": "Module 'openpyxl' missing.\npip install openpyxl",
        "pt_count":       "points",
        "statistics":     "Statistics",
        "mean":           "Mean (μ):",
        "std":            "Std. dev. (σ):",
        "minimum":        "Minimum:",
        "maximum":        "Maximum:",
        "peak_peak":      "Peak-Peak:",
        "count":          "Count:",
        "date_time":      "Date / Time:",
        "meas_function":  "Function:",
        "meas_range_hdr": "Range:",
        "resolution_hdr": "Resolution:",
        "interval_hdr":   "Interval (ms):",
        "num_points":     "No. of points:",
        "col_nr":         "#",
        "col_time":       "Time (s)",
        "col_abs":        "Clock",
        "chart_title":    "Measurement trend",
        "time_axis":      "Time (s)",
        "warn_stop_first": "Warning",
        "file_missing":   "Filename missing",
        "enter_filename":  "Please enter a filename.",
        "finalize_err":   "⚠ Finalisation failed: ",
        "file_open_err":  "⚠ Could not create file: ",
        "auto_save_info":   "🔄 Autosave at measurement end active",
        "excel_stats":  "Statistics in Excel",
        "excel_chart":  "Chart in Excel",
        "sheet_data":   "Data",
        "sheet_chart":  "Chart",
    },
}

# Device function labels per language
FUNC_LABELS = {
    "de": {
        "DC Voltage":    "DC Spannung",
        "AC Voltage":    "AC Spannung",
        "DC Current":    "DC Strom",
        "AC Current":    "AC Strom",
        "2W Resistance": "2W Widerstand",
        "4W Resistance": "4W Widerstand",
        "Frequency":     "Frequenz",
        "Period":        "Periode",
        "Continuity":    "Durchgang",
        "Diode":         "Diode",
    },
    "en": {
        "DC Voltage":    "DC Voltage",
        "AC Voltage":    "AC Voltage",
        "DC Current":    "DC Current",
        "AC Current":    "AC Current",
        "2W Resistance": "2W Resistance",
        "4W Resistance": "4W Resistance",
        "Frequency":     "Frequency",
        "Period":        "Period",
        "Continuity":    "Continuity",
        "Diode":         "Diode",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# Theme colour palettes
# ─────────────────────────────────────────────────────────────────────────────

THEMES: dict[str, dict[str, str]] = {
    "dark": {
        "bg":          "#1e1e2e",
        "bg2":         "#313244",
        "bg3":         "#45475a",
        "fg":          "#cdd6f4",
        "acc":         "#89b4fa",
        "acc2":        "#74c7ec",
        "acc3":        "#a6e3a1",
        "acc4":        "#cba6f7",
        "acc5":        "#fab387",
        "acc6":        "#f38ba8",
        "disp_bg":     "#11111b",
        "disp_val":    "#a6e3a1",
        "disp_unit":   "#89b4fa",
        "disp_func":   "#74c7ec",
        "disp_range":  "#cba6f7",
        "disp_sim":    "#fab387",
        "disp_label":  "#585b70",
        "plot_bg":     "#1e1e2e",
        "plot_axes":   "#313244",
        "plot_line":   "#89b4fa",
        "plot_grid":   "#45475a",
        "plot_text":   "#cdd6f4",
        "plot_stat":   "#a6e3a1",
        "plot_stat_bg":"#181825",
        "plot_stat_bd":"#45475a",
        "border":      "#45475a",
        "btn_bg":      "#313244",
        "btn_acc":     "#89b4fa",
        "btn_stop":    "#f38ba8",
        "btn_green":   "#a6e3a1",
        "entry_bg":    "#313244",
        "status_fg":   "#a6e3a1",
    },
    "bright": {
        "bg":          "#f0f0f0",
        "bg2":         "#dcdcdc",
        "bg3":         "#c8c8c8",
        "fg":          "#1a1a2e",
        "acc":         "#1565c0",
        "acc2":        "#0277bd",
        "acc3":        "#2e7d32",
        "acc4":        "#6a1b9a",
        "acc5":        "#e65100",
        "acc6":        "#c62828",
        "disp_bg":     "#ffffff",
        "disp_val":    "#1b5e20",
        "disp_unit":   "#0d47a1",
        "disp_func":   "#006064",
        "disp_range":  "#4a148c",
        "disp_sim":    "#bf360c",
        "disp_label":  "#9e9e9e",
        "plot_bg":     "#f5f5f5",
        "plot_axes":   "#e8e8e8",
        "plot_line":   "#1565c0",
        "plot_grid":   "#bdbdbd",
        "plot_text":   "#212121",
        "plot_stat":   "#1b5e20",
        "plot_stat_bg":"#ffffff",
        "plot_stat_bd":"#bdbdbd",
        "border":      "#bdbdbd",
        "btn_bg":      "#e0e0e0",
        "btn_acc":     "#1565c0",
        "btn_stop":    "#c62828",
        "btn_green":   "#2e7d32",
        "entry_bg":    "#ffffff",
        "status_fg":   "#1b5e20",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# Settings file path and defaults
# ─────────────────────────────────────────────────────────────────────────────

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

DEFAULT_SETTINGS = {
    "resource":    "SIMULATION",
    "function":    "DC Voltage",       # canonical English key
    "range":       "AUTO",
    "resolution":  "5½ Digit",
    "interval_ms": "500",
    "max_points":  "1000",
    "autoscale":    True,
    "scroll":       False,   # auto-scroll mode
    "scroll_win":   "30",    # visible time window in seconds
    "statistics":  True,
    "save_dir":    os.path.expanduser("~"),
    "prefix":      "Measurement",
    "suffix_mode": "datetime",
    "autosave":      True,
    "excel_stats":   True,   # include statistics block in Excel export
    "excel_chart":   True,   # include embedded chart in Excel export
    "language":      "de",
    "theme":         "dark",
}

def load_settings() -> dict:
    if os.path.isfile(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = DEFAULT_SETTINGS.copy()
            merged.update(data)
            return merged
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()

def save_settings(s: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARNING] Settings could not be saved: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# Device driver abstraction
# ─────────────────────────────────────────────────────────────────────────────

class Multimeter34401A:
    """Driver abstraction for HP/Agilent 34401A (GPIB, RS-232, USB-GPIB)."""

    # Canonical English function keys → SCPI command
    FUNCTIONS: dict[str, str] = {
        "DC Voltage":    "VOLT:DC",
        "AC Voltage":    "VOLT:AC",
        "DC Current":    "CURR:DC",
        "AC Current":    "CURR:AC",
        "2W Resistance": "RES",
        "4W Resistance": "FRES",
        "Frequency":     "FREQ",
        "Period":        "PER",
        "Continuity":    "CONT",
        "Diode":         "DIOD",
    }

    RANGES: dict[str, list[str]] = {
        "DC Voltage":    ["AUTO", "100 mV", "1 V", "10 V", "100 V", "1000 V"],
        "AC Voltage":    ["AUTO", "100 mV", "1 V", "10 V", "100 V", "750 V"],
        "DC Current":    ["AUTO", "10 mA", "100 mA", "1 A", "3 A"],
        "AC Current":    ["AUTO", "1 A", "3 A"],
        "2W Resistance": ["AUTO", "100 Ω", "1 kΩ", "10 kΩ", "100 kΩ", "1 MΩ", "10 MΩ", "100 MΩ"],
        "4W Resistance": ["AUTO", "100 Ω", "1 kΩ", "10 kΩ", "100 kΩ", "1 MΩ", "10 MΩ"],
        "Frequency":     ["AUTO"],
        "Period":        ["AUTO"],
        "Continuity":    ["–"],
        "Diode":         ["–"],
    }

    RESOLUTIONS: list[str] = ["3½ Digit", "4½ Digit", "5½ Digit", "6½ Digit"]
    NPLC_MAP:    dict[str, float] = {
        "3½ Digit": 0.02, "4½ Digit": 0.2, "5½ Digit": 1, "6½ Digit": 10
    }

    UNITS: dict[str, str] = {
        "DC Voltage": "V",  "AC Voltage": "V",
        "DC Current": "A",  "AC Current": "A",
        "2W Resistance": "Ω", "4W Resistance": "Ω",
        "Frequency": "Hz",  "Period": "s",
        "Continuity": "Ω",  "Diode": "V",
    }

    def __init__(self):
        self.instrument  = None
        self.simulation  = True
        self._sim_func   = "DC Voltage"

    # ── Connection ───────────────────────────────────────────────────────────

    def list_resources(self) -> list[str]:
        if not PYVISA_AVAILABLE:
            return []
        try:
            return list(pyvisa.ResourceManager().list_resources())
        except Exception:
            return []

    def connect(self, resource_string: str) -> bool:
        if not PYVISA_AVAILABLE or not resource_string:
            self.simulation = True
            return False
        try:
            rm = pyvisa.ResourceManager()
            self.instrument = rm.open_resource(resource_string)
            self.instrument.timeout = 5000
            idn = self.instrument.query("*IDN?")
            if "34401" not in idn and "34410" not in idn:
                raise ValueError(f"Unknown device: {idn}")
            self.instrument.write("*RST")
            self.instrument.write("*CLS")
            self.simulation = False
            return True
        except Exception as e:
            self.instrument = None
            self.simulation = True
            raise e

    def disconnect(self):
        if self.instrument:
            try:
                self.instrument.write("SYST:LOC")
                self.instrument.close()
            except Exception:
                pass
            self.instrument = None
        self.simulation = True

    # ── Configuration ────────────────────────────────────────────────────────

    def configure(self, function: str, range_str: str, resolution: str):
        """Configure the instrument for a measurement."""
        self._sim_func = function
        if self.simulation:
            return
        func_cmd  = self.FUNCTIONS.get(function, "VOLT:DC")
        nplc      = self.NPLC_MAP.get(resolution, 1)
        range_val = "DEF" if range_str in ("AUTO", "–") else self._parse_range(range_str)
        self.instrument.write(f"CONF:{func_cmd} {range_val}")
        if function not in ("Continuity", "Diode", "Frequency", "Period"):
            self.instrument.write(f"SENS:{func_cmd}:NPLC {nplc}")
        self.instrument.write("TRIG:SOUR IMM")
        self.instrument.write("TRIG:DEL:AUTO ON")
        self.instrument.write("SAMP:COUN 1")

    @staticmethod
    def _parse_range(s: str) -> str:
        """Convert human-readable range string to SCPI numeric value."""
        s = s.replace(" ", "").upper()
        for suffix, mult in sorted(
                {"MV": 1e-3, "V": 1, "MA": 1e-3, "A": 1,
                 "Ω": 1, "KΩ": 1e3, "MΩ": 1e6, "HZ": 1}.items(),
                key=lambda x: -len(x[0])):
            if s.endswith(suffix):
                try:
                    return str(float(s[:-len(suffix)]) * mult)
                except ValueError:
                    pass
        return "DEF"

    # ── Measurement ──────────────────────────────────────────────────────────

    def measure(self) -> float:
        if self.simulation:
            return self._simulate()
        try:
            return float(self.instrument.query("READ?").strip())
        except Exception:
            return float("nan")

    def _simulate(self) -> float:
        """Generate plausible simulated readings for the active function."""
        f = self._sim_func
        if "Voltage" in f:
            base = 5.0 if "DC" in f else 230.0
            return abs(base + random.gauss(0, base * 0.002))
        if "Current" in f:
            base = 0.1 if "DC" in f else 0.5
            return base + random.gauss(0, base * 0.005)
        if "Resistance" in f:
            return 1000.0 + random.gauss(0, 0.5)
        if "Frequency" in f:
            return 50.0 + random.gauss(0, 0.01)
        if "Period" in f:
            return 0.02 + random.gauss(0, 1e-6)
        return random.gauss(0, 0.001)


# ─────────────────────────────────────────────────────────────────────────────
# Main application
# ─────────────────────────────────────────────────────────────────────────────

class MultimeterApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.settings = load_settings()

        # Active language and theme
        self._lang  = self.settings.get("language", "de")
        self._theme = self.settings.get("theme", "dark")
        self.S      = STRINGS[self._lang]       # shortcut for string lookup
        self.C      = THEMES[self._theme]       # shortcut for colour lookup

        self.title(self.S["title"])
        self.geometry("1360x880")
        self.configure(bg=self.C["bg"])
        self.resizable(True, True)

        # Device driver
        self.dmm = Multimeter34401A()

        # Measurement data buffers
        self.data_timestamps: list[float] = []
        self.data_values:     list[float] = []
        self.running          = False
        self.measure_thread   = None
        self.data_queue       = queue.Queue()

        # Live saving state: CSV buffer during measurement, Excel on finalise
        self._autosave_counter        = 0
        self._autosave_path           = ""   # final .xlsx path
        self._autosave_csv_path       = ""   # temporary .csv buffer path
        self._autosave_csv_file       = None # open file handle
        self._autosave_csv_writer     = None # csv.writer instance
        self._autosave_t0             = 0.0
        self._autosave_abs_t0         = None
        self._autosave_data_row_count = 0
        self._autosave_lock           = threading.Lock()  # thread-safe csv writes
        # Legacy workbook refs kept for _finalize (now unused during measurement)
        self._autosave_wb             = None
        self._autosave_ws             = None
        self._autosave_row            = 0

        self._build_style()
        self._build_layout()
        self._apply_settings()
        self._update_plot_loop()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Trigger auto-connect after GUI is fully rendered
        self.after(200, self._auto_connect)

    # ── Shutdown ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._collect_settings()
        save_settings(self.settings)
        self.dmm.disconnect()
        self.destroy()

    # ── Settings helpers ─────────────────────────────────────────────────────

    def _collect_settings(self):
        """Read all UI variables into self.settings dict."""
        self.settings.update({
            "resource":    self.resource_var.get(),
            "function":    self._label_to_key(self.func_var.get()),
            "range":       self.range_var.get(),
            "resolution":  self.res_var.get(),
            "interval_ms": self.interval_var.get(),
            "max_points":  self.maxpts_var.get(),
            "autoscale":   bool(self.autoscale_var.get()),
            "scroll":      bool(self.scroll_var.get()),
            "scroll_win":  self.scroll_win_var.get(),
            "statistics":  bool(self.statistics_var.get()),
            "save_dir":    self.savedir_var.get(),
            "prefix":      self.prefix_var.get(),
            "suffix_mode": self.suffix_mode_var.get(),
            "autosave":    bool(self.autosave_var.get()),
            "excel_stats": bool(self.excel_stats_var.get()),
            "excel_chart": bool(self.excel_chart_var.get()),
            "language":    self._lang,
            "theme":       self._theme,
        })

    def _apply_settings(self):
        """Write self.settings into all UI variables."""
        s = self.settings
        self.resource_var.set(s["resource"])

        # Translate canonical function key to current language label
        func_label = FUNC_LABELS[self._lang].get(s.get("function", "DC Voltage"),
                                                  self._func_keys()[0])
        if func_label not in self._func_keys():
            func_label = self._func_keys()[0]
        self.func_var.set(func_label)

        self.interval_var.set(s["interval_ms"])
        self.maxpts_var.set(s["max_points"])
        self.autoscale_var.set(s["autoscale"])
        self.scroll_var.set(s.get("scroll", False))
        self.scroll_win_var.set(s.get("scroll_win", "30"))
        self._on_scroll_toggle()
        self.statistics_var.set(s["statistics"])
        self.savedir_var.set(s["save_dir"])
        self.prefix_var.set(s.get("prefix", "Measurement"))
        self.res_var.set(s["resolution"])
        self.autosave_var.set(s.get("autosave", True))
        self.excel_stats_var.set(s.get("excel_stats", True))
        self.excel_chart_var.set(s.get("excel_chart", True))
        self.suffix_mode_var.set(s.get("suffix_mode", "datetime"))
        self._on_function_change(_restore_range=s.get("range", "AUTO"))

    def _save_settings_now(self):
        self._collect_settings()
        save_settings(self.settings)
        self.status_var.set(self.S["settings_saved"] + SETTINGS_FILE)

    # ── Language / theme helpers ──────────────────────────────────────────────

    def _on_scroll_toggle(self):
        """Enable/disable the window-size entry based on scroll checkbox."""
        if hasattr(self, "scroll_win_entry"):
            state = "normal" if self.scroll_var.get() else "disabled"
            self.scroll_win_entry.configure(state=state)
            # Auto-scroll and autoscale are mutually exclusive
            if self.scroll_var.get():
                self.autoscale_var.set(False)

    def _t(self, key: str) -> str:
        """Translate a string key using the active language."""
        return self.S.get(key, key)

    def _func_keys(self) -> list[str]:
        """Return function display labels for the active language."""
        return list(FUNC_LABELS[self._lang].values())

    def _label_to_key(self, label: str) -> str:
        """Convert a display label back to the canonical English function key."""
        reverse = {v: k for k, v in FUNC_LABELS[self._lang].items()}
        return reverse.get(label, label)

    # ── Style ────────────────────────────────────────────────────────────────

    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        C = self.C
        bg, fg, acc = C["bg"], C["fg"], C["acc"]

        style.configure("TFrame",           background=bg)
        style.configure("TLabelframe",      background=bg, foreground=fg,
                        bordercolor=C["border"], relief="groove")
        style.configure("TLabelframe.Label", background=bg, foreground=acc,
                        font=("Segoe UI", 10, "bold"))
        style.configure("TLabel",           background=bg, foreground=fg,
                        font=("Segoe UI", 9))
        style.configure("TButton",          background=C["btn_bg"], foreground=fg,
                        font=("Segoe UI", 9, "bold"), padding=5)
        style.map("TButton",
                  background=[("active", C["bg2"]), ("pressed", C["bg3"])])
        style.configure("Accent.TButton",   background=C["btn_acc"], foreground=bg,
                        font=("Segoe UI", 9, "bold"), padding=5)
        style.map("Accent.TButton",
                  background=[("active", C["acc2"]), ("pressed", C["acc3"])])
        style.configure("Stop.TButton",     background=C["btn_stop"], foreground=bg,
                        font=("Segoe UI", 9, "bold"), padding=5)
        style.configure("Green.TButton",    background=C["btn_green"], foreground=bg,
                        font=("Segoe UI", 9, "bold"), padding=5)
        style.map("Green.TButton",
                  background=[("active", C["acc3"]), ("pressed", C["acc3"])])
        style.configure("TCombobox",        fieldbackground=C["entry_bg"],
                        background=C["entry_bg"], foreground=fg,
                        selectbackground=C["bg2"], selectforeground=fg)
        style.configure("TEntry",           fieldbackground=C["entry_bg"],
                        foreground=fg, insertcolor=fg)
        style.configure("TCheckbutton",     background=bg, foreground=fg,
                        font=("Segoe UI", 9))
        style.configure("TRadiobutton",     background=bg, foreground=fg,
                        font=("Segoe UI", 9))

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_layout(self):
        # Left sidebar
        sidebar = ttk.Frame(self, width=330)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=8)
        sidebar.pack_propagate(False)

        self._build_connection_frame(sidebar)
        self._build_function_frame(sidebar)
        self._build_acquisition_frame(sidebar)
        self._build_data_frame(sidebar)
        self._build_control_frame(sidebar)
        self._build_status_bar(sidebar)

        # Right area: display + plot
        right = ttk.Frame(self)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._build_display(right)
        self._build_plot(right)

    # ── Connection frame ─────────────────────────────────────────────────────

    def _build_connection_frame(self, parent):
        frm = ttk.LabelFrame(parent, text=f" {self._t('connection')} ", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))

        row1 = ttk.Frame(frm)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text=self._t("interface")).pack(side=tk.LEFT)

        self.resource_var   = tk.StringVar(value="SIMULATION")
        self.resource_combo = ttk.Combobox(row1, textvariable=self.resource_var,
                                           width=22, state="normal")
        self.resource_combo.pack(side=tk.LEFT, padx=(6, 4))
        ttk.Button(row1, text="🔍", width=3,
                   command=self._scan_resources).pack(side=tk.LEFT)

        row2 = ttk.Frame(frm)
        row2.pack(fill=tk.X, pady=(6, 0))
        self.btn_connect = ttk.Button(row2, text=self._t("connect"),
                                      style="Accent.TButton",
                                      command=self._toggle_connect)
        self.btn_connect.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.conn_status = ttk.Label(row2, text=self._t("disconnected"),
                                     foreground=self.C["btn_stop"], padding=(8, 0))
        self.conn_status.pack(side=tk.LEFT)

    # ── Device function frame ─────────────────────────────────────────────────

    def _build_function_frame(self, parent):
        frm = ttk.LabelFrame(parent, text=f" {self._t('functions')} ", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frm, text=self._t("meas_func")).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.func_var = tk.StringVar(value=self._func_keys()[0])
        func_cb = ttk.Combobox(frm, textvariable=self.func_var, width=20,
                                values=self._func_keys(), state="readonly")
        func_cb.grid(row=0, column=1, padx=(6, 0), pady=2, sticky=tk.EW)
        func_cb.bind("<<ComboboxSelected>>", self._on_function_change)

        ttk.Label(frm, text=self._t("meas_range")).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.range_var = tk.StringVar(value="AUTO")
        self.range_cb  = ttk.Combobox(frm, textvariable=self.range_var, width=20,
                                       state="readonly")
        self.range_cb.grid(row=1, column=1, padx=(6, 0), pady=2, sticky=tk.EW)
        self.range_cb.bind("<<ComboboxSelected>>", self._on_range_change)

        ttk.Label(frm, text=self._t("resolution")).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.res_var = tk.StringVar(value="5½ Digit")
        res_cb = ttk.Combobox(frm, textvariable=self.res_var, width=20,
                               values=Multimeter34401A.RESOLUTIONS, state="readonly")
        res_cb.grid(row=2, column=1, padx=(6, 0), pady=2, sticky=tk.EW)

        frm.columnconfigure(1, weight=1)
        # Initialise range list without touching display (not yet built)
        first_key = list(Multimeter34401A.RANGES.keys())[0]
        self.range_cb["values"] = Multimeter34401A.RANGES[first_key]
        self.range_var.set("AUTO")

    # ── Acquisition frame ─────────────────────────────────────────────────────

    def _build_acquisition_frame(self, parent):
        frm = ttk.LabelFrame(parent, text=f" {self._t('acquisition')} ", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frm, text=self._t("interval")).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.interval_var = tk.StringVar(value="500")
        ttk.Entry(frm, textvariable=self.interval_var, width=10).grid(
            row=0, column=1, padx=(6, 0), pady=2, sticky=tk.EW)

        ttk.Label(frm, text=self._t("max_points")).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.maxpts_var = tk.StringVar(value="1000")
        ttk.Entry(frm, textvariable=self.maxpts_var, width=10).grid(
            row=1, column=1, padx=(6, 0), pady=2, sticky=tk.EW)

        self.autoscale_var  = tk.BooleanVar(value=True)
        self.statistics_var = tk.BooleanVar(value=True)
        self.scroll_var     = tk.BooleanVar(value=False)
        self.scroll_win_var = tk.StringVar(value="30")
        ttk.Checkbutton(frm, text=self._t("autoscale"),
                        variable=self.autoscale_var).grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Checkbutton(frm, text=self._t("show_stats"),
                        variable=self.statistics_var).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=2)
        # Auto-scroll row
        self.scroll_cb = ttk.Checkbutton(
            frm, text=self._t("scroll"),
            variable=self.scroll_var,
            command=self._on_scroll_toggle)
        self.scroll_cb.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(frm, text=self._t("scroll_win")).grid(
            row=5, column=0, sticky=tk.W, pady=2)
        self.scroll_win_entry = ttk.Entry(
            frm, textvariable=self.scroll_win_var, width=10)
        self.scroll_win_entry.grid(
            row=5, column=1, padx=(6, 0), pady=2, sticky=tk.EW)
        frm.columnconfigure(1, weight=1)

    # ── File / autosave frame ─────────────────────────────────────────────────

    def _build_data_frame(self, parent):
        frm = ttk.LabelFrame(parent, text=f" {self._t('file')} ", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))

        # Save directory
        ttk.Label(frm, text=self._t("save_dir")).pack(anchor=tk.W)
        row_dir = ttk.Frame(frm)
        row_dir.pack(fill=tk.X, pady=(2, 6))
        self.savedir_var = tk.StringVar(value=os.path.expanduser("~"))
        ttk.Entry(row_dir, textvariable=self.savedir_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row_dir, text="📂", width=3,
                   command=self._browse_dir).pack(side=tk.LEFT, padx=(4, 0))

        # Separator
        tk.Frame(frm, bg=self.C["border"], height=1).pack(fill=tk.X, pady=(0, 6))

        # Prefix field
        ttk.Label(frm, text=self._t("prefix_lbl")).pack(anchor=tk.W)
        self.prefix_var = tk.StringVar(value="Measurement")
        ttk.Entry(frm, textvariable=self.prefix_var).pack(fill=tk.X, pady=(2, 6))

        # Suffix mode
        ttk.Label(frm, text=self._t("suffix_lbl")).pack(anchor=tk.W)
        self.suffix_mode_var = tk.StringVar(value="datetime")
        sf_row = ttk.Frame(frm)
        sf_row.pack(fill=tk.X, pady=(2, 4))
        ttk.Radiobutton(sf_row, text=self._t("suffix_dt"),
                        variable=self.suffix_mode_var, value="datetime",
                        command=self._update_preview).pack(side=tk.LEFT)
        ttk.Radiobutton(sf_row, text=self._t("suffix_cnt"),
                        variable=self.suffix_mode_var, value="counter",
                        command=self._update_preview).pack(side=tk.LEFT, padx=(10, 0))

        # Filename preview
        self.preview_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.preview_var,
                  foreground=self.C["acc2"], font=("Segoe UI", 7),
                  wraplength=310).pack(anchor=tk.W, pady=(0, 6))
        self.prefix_var.trace_add("write", lambda *_: self._update_preview())
        self._update_preview()

        # Separator
        tk.Frame(frm, bg=self.C["border"], height=1).pack(fill=tk.X, pady=(0, 6))

        # Autosave checkbox
        self.autosave_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text=self._t("autosave_lbl"),
                        variable=self.autosave_var).pack(anchor=tk.W, pady=(0, 2))

        # Excel content options
        self.excel_stats_var = tk.BooleanVar(value=True)
        self.excel_chart_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text=self._t("excel_stats"),
                        variable=self.excel_stats_var).pack(anchor=tk.W, pady=(0, 2))
        ttk.Checkbutton(frm, text=self._t("excel_chart"),
                        variable=self.excel_chart_var).pack(anchor=tk.W, pady=(0, 2))

        self.autosave_status_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.autosave_status_var,
                  foreground=self.C["acc3"], font=("Segoe UI", 7),
                  wraplength=310).pack(anchor=tk.W, pady=(0, 4))

        # Separator
        tk.Frame(frm, bg=self.C["border"], height=1).pack(fill=tk.X, pady=(0, 6))

        # Manual save + clear
        self.btn_save = ttk.Button(frm, text=self._t("save_now"),
                                   command=self._save_excel, state="disabled")
        self.btn_save.pack(fill=tk.X, pady=(0, 2))
        ttk.Button(frm, text=self._t("clear_data"),
                   command=self._clear_data).pack(fill=tk.X, pady=(0, 6))

        # Save settings button
        ttk.Button(frm, text=self._t("save_settings"),
                   style="Green.TButton",
                   command=self._save_settings_now).pack(fill=tk.X)

    # ── Control frame ─────────────────────────────────────────────────────────

    def _build_control_frame(self, parent):
        frm = ttk.Frame(parent)
        frm.pack(fill=tk.X, pady=(0, 6))

        # Gear icon opens settings dialog
        gear_row = ttk.Frame(frm)
        gear_row.pack(fill=tk.X, pady=(0, 4))
        self.btn_start = ttk.Button(gear_row, text=self._t("start"),
                                    style="Accent.TButton",
                                    command=self._start_measurement)
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(gear_row, text="⚙", width=4,
                   command=self._open_settings_dialog).pack(side=tk.LEFT)

        self.btn_stop = ttk.Button(frm, text=self._t("stop"),
                                   style="Stop.TButton",
                                   command=self._stop_measurement,
                                   state="disabled")
        self.btn_stop.pack(fill=tk.X)

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_status_bar(self, parent):
        self.status_var = tk.StringVar(value=self._t("ready"))
        ttk.Label(parent, textvariable=self.status_var,
                  foreground=self.C["status_fg"], wraplength=315,
                  font=("Segoe UI", 8)).pack(anchor=tk.W, pady=(6, 0))

    # ── Large numeric display ─────────────────────────────────────────────────

    def _build_display(self, parent):
        disp_frame = tk.Frame(parent, bg=self.C["bg"], height=125)
        disp_frame.pack(fill=tk.X, pady=(0, 6))
        disp_frame.pack_propagate(False)

        inner = tk.Frame(disp_frame, bg=self.C["disp_bg"], relief="sunken", bd=2)
        inner.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        self.display_value = tk.StringVar(value="- - - - - -")
        self.display_unit  = tk.StringVar(value="V")
        self.display_func  = tk.StringVar(value=self._func_keys()[0])
        self.display_range = tk.StringVar(value="AUTO")

        # Top info row: function | range label | range | SIM
        top = tk.Frame(inner, bg=self.C["disp_bg"])
        top.pack(fill=tk.X, padx=12, pady=(6, 0))

        tk.Label(top, textvariable=self.display_func,
                 bg=self.C["disp_bg"], fg=self.C["disp_func"],
                 font=("Courier New", 12, "bold")).pack(side=tk.LEFT)

        self.sim_indicator = tk.Label(top, text="SIM",
                                      bg=self.C["disp_bg"], fg=self.C["disp_sim"],
                                      font=("Courier New", 9))
        self.sim_indicator.pack(side=tk.RIGHT)

        tk.Label(top, textvariable=self.display_range,
                 bg=self.C["disp_bg"], fg=self.C["disp_range"],
                 font=("Courier New", 12, "bold")).pack(side=tk.RIGHT, padx=(0, 10))

        tk.Label(top, text=self._t("range_lbl"),
                 bg=self.C["disp_bg"], fg=self.C["disp_label"],
                 font=("Courier New", 9)).pack(side=tk.RIGHT)

        # Value row
        val_row = tk.Frame(inner, bg=self.C["disp_bg"])
        val_row.pack(fill=tk.X, padx=12, pady=(2, 0))

        tk.Label(val_row, textvariable=self.display_value,
                 bg=self.C["disp_bg"], fg=self.C["disp_val"],
                 font=("Courier New", 38, "bold")).pack(side=tk.LEFT)
        tk.Label(val_row, textvariable=self.display_unit,
                 bg=self.C["disp_bg"], fg=self.C["disp_unit"],
                 font=("Courier New", 38, "bold")).pack(side=tk.LEFT, padx=(6, 0))

    # ── Live plot ─────────────────────────────────────────────────────────────

    def _build_plot(self, parent):
        C = self.C
        plot_frame = ttk.Frame(parent)
        plot_frame.pack(fill=tk.BOTH, expand=True)

        self.fig = Figure(figsize=(8, 4.5), facecolor=C["plot_bg"])
        self.ax  = self.fig.add_subplot(111)
        self._style_axes()
        self.line, = self.ax.plot([], [], color=C["plot_line"], linewidth=1.5,
                                  antialiased=True)
        self.fig.tight_layout(pad=1.2)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        tb_frame = ttk.Frame(plot_frame)
        tb_frame.pack(fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(self.canvas, tb_frame)
        self.toolbar.update()

        # Statistics text overlay
        self.stat_text = self.ax.text(
            0.01, 0.98, "", transform=self.ax.transAxes,
            verticalalignment="top", fontfamily="monospace", fontsize=8,
            color=C["plot_stat"],
            bbox=dict(boxstyle="round,pad=0.3", facecolor=C["plot_stat_bg"],
                      edgecolor=C["plot_stat_bd"], alpha=0.85))

    def _style_axes(self):
        C = self.C
        self.ax.set_facecolor(C["plot_axes"])
        self.ax.tick_params(colors=C["plot_text"], labelsize=8)
        self.ax.xaxis.label.set_color(C["plot_text"])
        self.ax.yaxis.label.set_color(C["plot_text"])
        for spine in self.ax.spines.values():
            spine.set_edgecolor(C["plot_grid"])
        self.ax.grid(True, color=C["plot_grid"], linewidth=0.5, alpha=0.6)
        self.ax.set_xlabel(self._t("time_axis"), color=C["plot_text"], fontsize=9)
        self.ax.set_ylabel("Value", color=C["plot_text"], fontsize=9)
        self.ax.set_title(self._t("chart_title"), color=C["plot_text"], fontsize=10)

    # ── Settings dialog (gear) ────────────────────────────────────────────────

    def _open_settings_dialog(self):
        """Open a modal dialog for language and theme selection."""
        dlg = tk.Toplevel(self)
        dlg.title(self._t("settings_title"))
        dlg.configure(bg=self.C["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        pad = dict(padx=12, pady=6)
        # Language
        ttk.Label(dlg, text=self._t("lang_lbl")).grid(row=0, column=0, sticky=tk.W, **pad)
        lang_var = tk.StringVar(value=self._lang)
        lang_frame = ttk.Frame(dlg)
        lang_frame.grid(row=0, column=1, sticky=tk.W, **pad)
        ttk.Radiobutton(lang_frame, text="Deutsch", variable=lang_var,
                        value="de").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(lang_frame, text="English", variable=lang_var,
                        value="en").pack(side=tk.LEFT)

        # Theme
        ttk.Label(dlg, text=self._t("theme_lbl")).grid(row=1, column=0, sticky=tk.W, **pad)
        theme_var = tk.StringVar(value=self._theme)
        theme_frame = ttk.Frame(dlg)
        theme_frame.grid(row=1, column=1, sticky=tk.W, **pad)
        ttk.Radiobutton(theme_frame, text=self._t("theme_dark"),
                        variable=theme_var, value="dark").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(theme_frame, text=self._t("theme_bright"),
                        variable=theme_var, value="bright").pack(side=tk.LEFT)

        def apply_and_restart():
            self._collect_settings()
            self.settings["language"] = lang_var.get()
            self.settings["theme"]    = theme_var.get()
            save_settings(self.settings)
            dlg.destroy()
            self.destroy()
            # Restart application in-process
            app = MultimeterApp()
            app.mainloop()

        ttk.Button(dlg, text=self._t("apply"), style="Accent.TButton",
                   command=apply_and_restart).grid(
            row=2, column=0, columnspan=2, pady=12, padx=12, sticky=tk.EW)

    # ── Callbacks ────────────────────────────────────────────────────────────

    def _on_function_change(self, event=None, _restore_range: str = None):
        """Update range list and display when measurement function changes."""
        label = self.func_var.get()
        key   = self._label_to_key(label)
        ranges = Multimeter34401A.RANGES.get(key, ["AUTO"])
        self.range_cb["values"] = ranges

        if _restore_range and _restore_range in ranges:
            self.range_var.set(_restore_range)
        else:
            self.range_var.set(ranges[0])

        unit = Multimeter34401A.UNITS.get(key, "")
        if hasattr(self, "display_unit"):
            self.display_unit.set(unit)
            self.display_func.set(label)
            self.display_range.set(self.range_var.get())
        if hasattr(self, "ax"):
            self.ax.set_ylabel(f"{label} ({unit})",
                               color=self.C["plot_text"], fontsize=9)
        if hasattr(self, "canvas"):
            self.canvas.draw_idle()

    def _on_range_change(self, event=None):
        if hasattr(self, "display_range"):
            self.display_range.set(self.range_var.get())

    def _scan_resources(self):
        resources = self.dmm.list_resources()
        self.resource_combo["values"] = ["SIMULATION"] + resources
        if resources:
            self.status_var.set(f"{len(resources)} {self._t('devices_found')}")
        else:
            self.status_var.set(self._t("no_devices"))

    def _toggle_connect(self):
        if not self.dmm.simulation:
            self.dmm.disconnect()
            self.conn_status.configure(text=self._t("disconnected"),
                                       foreground=self.C["btn_stop"])
            self.btn_connect.configure(text=self._t("connect"), style="Accent.TButton")
            self.sim_indicator.configure(text="SIM")
            self.status_var.set(self._t("disconnected"))
        else:
            res = self.resource_var.get()
            if res == "SIMULATION":
                self.conn_status.configure(text=self._t("simulation"),
                                           foreground=self.C["acc5"])
                self.sim_indicator.configure(text="SIM")
                self.status_var.set(self._t("simulation"))
            else:
                try:
                    self.dmm.connect(res)
                    self.conn_status.configure(text=self._t("connected"),
                                               foreground=self.C["acc3"])
                    self.btn_connect.configure(text=self._t("disconnect"),
                                               style="Stop.TButton")
                    self.sim_indicator.configure(text="")
                    self.status_var.set(f"{self._t('connected')}: {res}")
                except Exception as e:
                    messagebox.showerror("Error", str(e))
                    self.status_var.set(f"Error: {e}")

    # ── Auto-connect ──────────────────────────────────────────────────────────

    def _auto_connect(self):
        """Attempt to connect to the saved resource on startup (background)."""
        res = self.resource_var.get()
        if not res or res == "SIMULATION" or not PYVISA_AVAILABLE:
            return
        self.status_var.set(f"{self._t('connecting')} {res} …")
        self.btn_connect.configure(state="disabled")
        threading.Thread(target=self._auto_connect_thread,
                         args=(res,), daemon=True).start()

    def _auto_connect_thread(self, res: str):
        try:
            self.dmm.connect(res)
            self.after(0, self._on_auto_connect_success, res)
        except Exception:
            self.after(0, self._on_auto_connect_fail)

    def _on_auto_connect_success(self, res: str):
        self.conn_status.configure(text=self._t("connected"),
                                   foreground=self.C["acc3"])
        self.btn_connect.configure(text=self._t("disconnect"),
                                   style="Stop.TButton", state="normal")
        self.sim_indicator.configure(text="")
        self.status_var.set(f"{self._t('auto_connected')} {res}")

    def _on_auto_connect_fail(self):
        self.btn_connect.configure(state="normal")
        self.status_var.set(self._t("auto_conn_fail"))

    # ── Measurement control ───────────────────────────────────────────────────

    def _start_measurement(self):
        try:
            interval_ms = int(self.interval_var.get())
            if interval_ms < 50:
                raise ValueError(self._t("min_interval"))
        except ValueError as e:
            messagebox.showerror(self._t("input_error"), str(e))
            return

        func  = self.func_var.get()
        key   = self._label_to_key(func)
        unit  = Multimeter34401A.UNITS.get(key, "")
        self.dmm.configure(key, self.range_var.get(), self.res_var.get())

        # Clear previous plot data before starting a new measurement
        self.data_timestamps.clear()
        self.data_values.clear()
        self.line.set_data([], [])
        self.stat_text.set_text("")
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas.draw_idle()
        self.display_value.set("- - - - - -")

        # Update display immediately
        self.display_func.set(func)
        self.display_unit.set(unit)
        self.display_range.set(self.range_var.get())

        self.running = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_save.configure(state="disabled")
        self.autosave_status_var.set("")

        # Open CSV buffer for live saving (converted to Excel at measurement end)
        if self.autosave_var.get() and OPENPYXL_AVAILABLE:
            self._open_live_workbook()

        self.status_var.set(
            f"{self._t('meas_running')}  {func}  {self.range_var.get()}")

        self.measure_thread = threading.Thread(
            target=self._measure_loop, args=(interval_ms,), daemon=True)
        self.measure_thread.start()

    def _stop_measurement(self):
        self.running = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        if self.data_values:
            self.btn_save.configure(state="normal")
        self._finalize_live_workbook()
        n = len(self.data_values)
        self.status_var.set(
            f"{self._t('meas_stopped')} – {n} {self._t('pts_recorded')}")

    def _on_maxpts_reached(self):
        """Called in main thread when max-points limit is hit."""
        self.running = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        if self.data_values:
            self.btn_save.configure(state="normal")
        self._finalize_live_workbook()
        n = self.maxpts_var.get()
        self.status_var.set(
            f"{self._t('max_pts_reached')} ({n}) – "
            f"{len(self.data_values)} {self._t('pts_recorded')}")

    def _measure_loop(self, interval_ms: int):
        """Background thread: acquire measurements and push to queue."""
        t0 = time.time()
        while self.running:
            t_start = time.time()
            val     = self.dmm.measure()
            elapsed = time.time() - t0
            self.data_queue.put((elapsed, val))

            # Append to CSV buffer immediately (non-blocking)
            if self._autosave_csv_writer is not None:
                self._append_live_row(elapsed, val)

            # Stop when max points reached
            try:
                maxpts = int(self.maxpts_var.get())
            except ValueError:
                maxpts = 0
            if maxpts > 0 and (len(self.data_timestamps) + self.data_queue.qsize()) >= maxpts:
                self.after(0, self._on_maxpts_reached)
                break

            sleep_time = (interval_ms / 1000.0) - (time.time() - t_start)
            if sleep_time > 0:
                time.sleep(sleep_time)

    # ── Plot update loop (main thread) ────────────────────────────────────────

    def _update_plot_loop(self):
        changed = False
        while not self.data_queue.empty():
            t, v = self.data_queue.get_nowait()
            self.data_timestamps.append(t)
            self.data_values.append(v)
            changed = True

        if changed and self.data_values:
            key  = self._label_to_key(self.func_var.get())
            unit = Multimeter34401A.UNITS.get(key, "")
            self.display_value.set(f"{self.data_values[-1]:>+14.7g}")

            self.line.set_data(self.data_timestamps, self.data_values)

            if self.scroll_var.get():
                # Auto-scroll: fixed time window, newest data always visible
                try:
                    win = float(self.scroll_win_var.get())
                except ValueError:
                    win = 30.0
                tmax = self.data_timestamps[-1]
                tmin = max(self.data_timestamps[0], tmax - win)
                # Y-range from visible points only
                visible = [v for t, v in zip(self.data_timestamps,
                                             self.data_values)
                           if tmin <= t <= tmax]
                if visible:
                    vmin, vmax = min(visible), max(visible)
                    margin = ((vmax - vmin) * 0.05
                              if vmax != vmin else abs(vmax) * 0.05 + 1e-9)
                    self.ax.set_ylim(vmin - margin, vmax + margin)
                self.ax.set_xlim(tmin, tmax if tmax > tmin else tmin + 1)
            elif self.autoscale_var.get():
                self.ax.relim()
                self.ax.autoscale_view()
            else:
                tmin = min(self.data_timestamps)
                tmax = max(self.data_timestamps)
                vmin = min(self.data_values)
                vmax = max(self.data_values)
                margin = (vmax - vmin) * 0.05 if vmax != vmin else abs(vmax) * 0.05 + 1e-9
                self.ax.set_xlim(tmin, tmax if tmax > tmin else tmin + 1)
                self.ax.set_ylim(vmin - margin, vmax + margin)

            if self.statistics_var.get() and len(self.data_values) >= 2:
                arr = np.array(self.data_values)
                txt = (f"n   = {len(arr)}\n"
                       f"μ   = {arr.mean():.6g} {unit}\n"
                       f"σ   = {arr.std():.4g} {unit}\n"
                       f"min = {arr.min():.6g} {unit}\n"
                       f"max = {arr.max():.6g} {unit}")
                self.stat_text.set_text(txt)
            else:
                self.stat_text.set_text("")

            self.canvas.draw_idle()

            # Enable save button as soon as measurement stopped and data exists
            if not self.running:
                self.btn_save.configure(state="normal")

        self.after(80, self._update_plot_loop)

    # ── Filename / suffix helpers ─────────────────────────────────────────────

    def _build_suffix(self) -> str:
        """Generate filename suffix according to selected mode."""
        if self.suffix_mode_var.get() == "datetime":
            return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        else:
            self._autosave_counter += 1
            return f"{self._autosave_counter:04d}"

    def _build_save_filename(self) -> str:
        """Build full filename: <prefix>_<suffix>.xlsx"""
        prefix = self.prefix_var.get().strip() or "Measurement"
        return f"{prefix}_{self._build_suffix()}.xlsx"

    def _update_preview(self):
        """Refresh the filename preview label."""
        if not hasattr(self, "preview_var"):
            return
        prefix = self.prefix_var.get().strip() if hasattr(self, "prefix_var") else "M"
        if not prefix:
            prefix = "Measurement"
        mode   = self.suffix_mode_var.get() if hasattr(self, "suffix_mode_var") else "datetime"
        suffix = (datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                  if mode == "datetime" else f"{self._autosave_counter + 1:04d}")
        self.preview_var.set(f"{self._t('preview')} {prefix}_{suffix}.xlsx")

    # ── Live workbook (per-value saving) ──────────────────────────────────────

    def _open_live_workbook(self):
        """Open a lightweight CSV buffer at measurement start.
        The CSV is converted to Excel only once when the measurement ends.
        This avoids re-saving the entire workbook on every sample.
        """
        if not OPENPYXL_AVAILABLE:
            return
        try:
            os.makedirs(self.savedir_var.get(), exist_ok=True)
            filename              = self._build_save_filename()
            self._autosave_path   = os.path.join(self.savedir_var.get(), filename)
            # Temporary CSV sits next to the final xlsx
            self._autosave_csv_path = self._autosave_path.replace(".xlsx", ".csv~")
            self._autosave_t0     = time.time()
            self._autosave_abs_t0 = datetime.datetime.now()
            self._autosave_data_row_count = 0

            key  = self._label_to_key(self.func_var.get())
            func = self.func_var.get()
            unit = Multimeter34401A.UNITS.get(key, "")

            # Open CSV for streaming append – newline="" required by csv module
            self._autosave_csv_file   = open(
                self._autosave_csv_path, "w", newline="", encoding="utf-8")
            self._autosave_csv_writer = csv.writer(self._autosave_csv_file)

            # Write header row
            self._autosave_csv_writer.writerow([
                self._t("col_nr"),
                self._t("col_time"),
                f"{func} ({unit})",
                self._t("col_abs"),
            ])
            self._autosave_csv_file.flush()

            self.after(0, lambda fn=filename: self.autosave_status_var.set(
                self._t("live_active") + fn))
        except Exception as e:
            self._autosave_csv_file   = None
            self._autosave_csv_writer = None
            self.after(0, lambda err=e: self.autosave_status_var.set(
                self._t("file_open_err") + str(err)))

    def _append_live_row(self, elapsed: float, val: float):
        """Append one measurement row to the CSV buffer.
        Writing a single CSV line is ~1000x faster than re-saving a workbook,
        ensuring the measurement interval is never delayed by file I/O.
        """
        if self._autosave_csv_writer is None:
            return
        try:
            with self._autosave_lock:
                self._autosave_data_row_count += 1
                i        = self._autosave_data_row_count
                abs_time = self._autosave_abs_t0 + datetime.timedelta(seconds=elapsed)
                self._autosave_csv_writer.writerow([
                    i,
                    round(elapsed, 4),
                    round(val, 9),
                    abs_time.strftime("%H:%M:%S.%f")[:-3],
                ])
                self._autosave_csv_file.flush()  # OS-level flush – data on disk immediately
        except Exception:
            pass  # Silent – never interrupt the measurement thread

    def _finalize_live_workbook(self):
        """Convert the CSV buffer to a proper Excel file in a background thread.
        Steps:
          1. Close the CSV file handle
          2. Read all rows from CSV
          3. Build workbook with plain table + optional stats + optional chart
          4. Delete the temporary CSV
        """
        if self._autosave_csv_writer is None and self._autosave_data_row_count == 0:
            return
        # Snapshot values needed in the background thread
        n          = self._autosave_data_row_count
        csv_path   = self._autosave_csv_path
        xlsx_path  = self._autosave_path
        timestamps = list(self.data_timestamps)
        values     = list(self.data_values)
        inc_stats  = self.excel_stats_var.get()
        inc_chart  = self.excel_chart_var.get()
        key        = self._label_to_key(self.func_var.get())
        func       = self.func_var.get()
        unit       = Multimeter34401A.UNITS.get(key, "")
        strings    = dict(self.S)   # language strings snapshot

        # Close CSV file before reading it in the background thread
        try:
            if self._autosave_csv_file:
                self._autosave_csv_file.close()
        except Exception:
            pass
        finally:
            self._autosave_csv_file   = None
            self._autosave_csv_writer = None
            self._autosave_data_row_count = 0

        # Convert in background so GUI stays responsive
        t = threading.Thread(
            target=self._csv_to_excel_thread,
            args=(csv_path, xlsx_path, n, timestamps, values,
                  inc_stats, inc_chart, func, unit, strings),
            daemon=True)
        t.start()

    def _csv_to_excel_thread(self, csv_path, xlsx_path, n,
                              timestamps, values,
                              inc_stats, inc_chart,
                              func, unit, strings):
        """Background thread: read CSV buffer and write final Excel file."""
        def _t(k): return strings.get(k, k)
        try:
            # ── Read CSV rows ─────────────────────────────────────────────────
            rows = []
            if os.path.isfile(csv_path):
                with open(csv_path, "r", newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    next(reader, None)   # skip header (we write our own)
                    for row in reader:
                        rows.append(row)

            # ── Build workbook ────────────────────────────────────────────────
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = _t("sheet_data")

            bold = Font(bold=True)
            # Header row
            for col, txt in enumerate([
                    _t("col_nr"), _t("col_time"),
                    f"{func} ({unit})", _t("col_abs")], start=1):
                ws.cell(row=1, column=col, value=txt).font = bold

            # Data rows from CSV
            for i, row in enumerate(rows, start=1):
                xl_row = 1 + i
                try:
                    ws.cell(row=xl_row, column=1, value=int(row[0]))
                    ws.cell(row=xl_row, column=2, value=float(row[1]))
                    ws.cell(row=xl_row, column=3, value=float(row[2]))
                    ws.cell(row=xl_row, column=4, value=row[3])
                except (IndexError, ValueError):
                    pass   # skip malformed rows

            # Auto column widths
            for col_letter, width in zip("ABCD", [8, 14, 20, 16]):
                ws.column_dimensions[col_letter].width = width

            # ── Optional statistics ───────────────────────────────────────────
            actual_n = len(rows)
            if inc_stats and actual_n > 0:
                # Extract values from CSV rows (column index 2)
                try:
                    csv_vals = np.array([float(r[2]) for r in rows])
                except (IndexError, ValueError):
                    csv_vals = np.array([])
                if csv_vals.size > 0:
                    stat_row = 1 + actual_n + 2
                    ws.cell(row=stat_row, column=1,
                            value=_t("statistics")).font = bold
                    for j, (k, v) in enumerate([
                        (_t("count"),     str(actual_n)),
                        (_t("mean"),      f"{csv_vals.mean():.9g} {unit}"),
                        (_t("std"),       f"{csv_vals.std():.6g} {unit}"),
                        (_t("minimum"),   f"{csv_vals.min():.9g} {unit}"),
                        (_t("maximum"),   f"{csv_vals.max():.9g} {unit}"),
                        (_t("peak_peak"), f"{csv_vals.max() - csv_vals.min():.6g} {unit}"),
                    ], start=1):
                        ws.cell(row=stat_row + j, column=1, value=k).font = bold
                        ws.cell(row=stat_row + j, column=2, value=v)

            # ── Optional chart on separate sheet ──────────────────────────────
            # Use only the CSV rows (= this measurement only) as chart source.
            # Never use the global timestamps/values lists which may contain
            # data from previous measurement runs.
            if inc_chart and actual_n > 0 and rows:
                ws_c = wb.create_sheet(_t("sheet_chart"))
                ws_c.cell(row=1, column=1, value=_t("col_time")).font = bold
                ws_c.cell(row=1, column=2, value=f"{func} ({unit})").font = bold
                # Normalise time to start at 0
                try:
                    t0_csv = float(rows[0][1])
                except (IndexError, ValueError):
                    t0_csv = 0.0
                for idx, row in enumerate(rows, start=2):
                    try:
                        ws_c.cell(row=idx, column=1,
                                  value=round(float(row[1]) - t0_csv, 4))
                        ws_c.cell(row=idx, column=2,
                                  value=round(float(row[2]), 9))
                    except (IndexError, ValueError):
                        pass

                chart              = LineChart()
                chart.title        = f"{func} – {_t('chart_title')}"
                chart.style        = 10
                chart.y_axis.title = f"{func} ({unit})"
                chart.x_axis.title = _t("time_axis")
                data_ref = Reference(ws_c, min_col=2, min_row=1,
                                     max_row=actual_n + 1)
                time_ref = Reference(ws_c, min_col=1, min_row=2,
                                     max_row=actual_n + 1)
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(time_ref)
                chart.width  = 22
                chart.height = 14
                ws_c.add_chart(chart, "D2")

            wb.save(xlsx_path)
            fname = os.path.basename(xlsx_path)
            self.after(0, lambda fn=fname, cnt=actual_n:
                       self.autosave_status_var.set(
                           f"{_t('live_done')} ({cnt} {_t('pt_count')})  →  {fn}"))
        except Exception as e:
            self.after(0, lambda err=e: self.autosave_status_var.set(
                _t("finalize_err") + str(err)))
        finally:
            # Remove temporary CSV buffer
            try:
                if os.path.isfile(csv_path):
                    os.remove(csv_path)
            except Exception:
                pass

    # ── File dialogs ──────────────────────────────────────────────────────────

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.savedir_var.get(),
                                    title=self._t("save_dir"))
        if d:
            self.savedir_var.set(d)

    def _clear_data(self):
        if self.running:
            messagebox.showwarning(self._t("warn_stop_first"),
                                   self._t("stop_first"))
            return
        self.data_timestamps.clear()
        self.data_values.clear()
        self.line.set_data([], [])
        self.ax.relim()
        self.ax.autoscale_view()
        self.stat_text.set_text("")
        self.canvas.draw_idle()
        self.display_value.set("- - - - - -")
        self.btn_save.configure(state="disabled")
        self.status_var.set(self._t("data_cleared"))

    # ── Manual Excel export ───────────────────────────────────────────────────

    def _save_excel(self):
        if not self.data_values:
            messagebox.showinfo("Info", self._t("no_data"))
            return
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Error", self._t("openpyxl_missing"))
            return
        filename = self._build_save_filename()
        os.makedirs(self.savedir_var.get(), exist_ok=True)
        path = os.path.join(self.savedir_var.get(), filename)
        try:
            self._write_excel_data(path, list(self.data_timestamps),
                                   list(self.data_values))
            messagebox.showinfo("OK", self._t("saved") + path)
            self.status_var.set(self._t("saved") + path)
        except Exception as e:
            messagebox.showerror(self._t("save_error"), str(e))

    def _write_excel_data(self, path: str, timestamps: list, values: list):
        """Write plain measurement table; optional stats/chart controlled by checkboxes."""
        key  = self._label_to_key(self.func_var.get())
        func = self.func_var.get()
        unit = Multimeter34401A.UNITS.get(key, "")
        now  = datetime.datetime.now()
        t0   = timestamps[0]
        abs_t0 = now - datetime.timedelta(seconds=timestamps[-1])

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self._t("sheet_data")

        # ── Plain header row (bold, no fill) ──────────────────────────────────
        bold = Font(bold=True)
        for col, txt in enumerate([
                self._t("col_nr"),
                self._t("col_time"),
                f"{func} ({unit})",
                self._t("col_abs")], start=1):
            ws.cell(row=1, column=col, value=txt).font = bold

        # ── Data rows – no fill, no borders ───────────────────────────────────
        for i, (ts, val) in enumerate(zip(timestamps, values), start=1):
            row      = 1 + i
            abs_time = abs_t0 + datetime.timedelta(seconds=ts)
            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=round(ts - t0, 4))
            ws.cell(row=row, column=3, value=round(val, 9))
            ws.cell(row=row, column=4, value=abs_time.strftime("%H:%M:%S.%f")[:-3])

        # ── Optional statistics block ─────────────────────────────────────────
        if self.excel_stats_var.get():
            arr      = np.array(values)
            stat_row = 1 + len(values) + 2   # two blank rows gap
            ws.cell(row=stat_row, column=1,
                    value=self._t("statistics")).font = bold
            for j, (k, v) in enumerate([
                (self._t("count"),     str(len(values))),
                (self._t("mean"),      f"{arr.mean():.9g} {unit}"),
                (self._t("std"),       f"{arr.std():.6g} {unit}"),
                (self._t("minimum"),   f"{arr.min():.9g} {unit}"),
                (self._t("maximum"),   f"{arr.max():.9g} {unit}"),
                (self._t("peak_peak"), f"{arr.max() - arr.min():.6g} {unit}"),
            ], start=1):
                ws.cell(row=stat_row + j, column=1, value=k).font = bold
                ws.cell(row=stat_row + j, column=2, value=v)

        # ── Optional embedded chart on a separate sheet ───────────────────────
        if self.excel_chart_var.get():
            ws_chart = wb.create_sheet(self._t("sheet_chart"))
            ws_chart.cell(row=1, column=1,
                          value=self._t("col_time")).font = bold
            ws_chart.cell(row=1, column=2,
                          value=f"{func} ({unit})").font = bold
            for idx, (ts, val) in enumerate(zip(timestamps, values), start=2):
                ws_chart.cell(row=idx, column=1, value=round(ts - t0, 4))
                ws_chart.cell(row=idx, column=2, value=round(val, 9))

            chart              = LineChart()
            chart.title        = f"{func} – {self._t('chart_title')}"
            chart.style        = 10
            chart.y_axis.title = f"{func} ({unit})"
            chart.x_axis.title = self._t("time_axis")
            data_ref = Reference(ws_chart, min_col=2, min_row=1,
                                 max_row=len(values) + 1)
            time_ref = Reference(ws_chart, min_col=1, min_row=2,
                                 max_row=len(values) + 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(time_ref)
            chart.width  = 22
            chart.height = 14
            ws_chart.add_chart(chart, "D2")

        # ── Auto-fit column widths on data sheet ──────────────────────────────
        for col_cells in ws.iter_rows():
            for cell in col_cells:
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                cl = get_column_letter(cell.column)
                w  = len(str(cell.value)) + 4
                current = ws.column_dimensions[cl].width or 0
                if w > current:
                    ws.column_dimensions[cl].width = min(w, 45)

        wb.save(path)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    missing = []
    if not PYVISA_AVAILABLE:
        missing.append("pyvisa")
    if not OPENPYXL_AVAILABLE:
        missing.append("openpyxl")
    if missing:
        print(f"[WARNING] Missing modules: {', '.join(missing)}")
        print(f"Install with: pip install {' '.join(missing)}")
        print("Application will run in limited mode.\n")

    app = MultimeterApp()
    app.mainloop()
