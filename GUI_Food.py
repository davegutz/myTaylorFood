#! /usr/bin/env python3
# GUI_Food.py - Food Expense & Photo Analysis Tracking GUI
# In the style of mySOC/SOC_Particle/pyStateOfCharge (Dave Gutz)
#
# Copyright (C) 2026 Dave Gutz
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation;
# version 2.1 of the License.

import os
import sys
import platform
import csv
import json
import time
import datetime
import shutil
import re
from pathlib import Path, PurePosixPath
from configparser import ConfigParser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Optional PIL support for robust image handling
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Optional RapidOCR support for reading receipt text
try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_OCR = True
    ocr_engine = RapidOCR()
except ImportError:
    HAS_OCR = False
    ocr_engine = None

# Optional openpyxl for Excel workbook handling
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
  
plat = sys.platform

# Platform-specific paths
if plat == "linux":
    default_base_dir = "/home/daveg/gdrive/0 - Taylor/myTaylorFood"
    default_log_dir = os.path.expanduser("~/.local")
elif plat == "darwin":
    default_base_dir = "/Users/daveg/Library/CloudStorage/GoogleDrive-davegutz2006@gmail.com/My Drive/0 - Taylor/myTaylorFood"
    default_log_dir = os.path.expanduser("~/.local")
else:
    local_app_data = os.getenv("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    default_base_dir = "G:/My Drive/0 - Taylor/myTaylorFood"
    default_log_dir = str(Path(local_app_data) / "Temp")

default_archive_dir = os.path.join(default_base_dir, "Archive")
default_excel_path = os.path.join(default_base_dir, "TaylorMealRecords.xlsx")
default_csv_path = os.path.join(default_base_dir, "TaylorMealRecords.csv")

# Configuration default values
default_dict = {
    "paths": {
        "base_folder": default_base_dir,
        "archive_folder": default_archive_dir,
        "excel_path": default_excel_path,
        "last_photo_folder": default_base_dir,
        "last_photo": "",
    },
    "options": {
        "meal_type": "Breakfast",
        "auto_analyze": "False",
        "confirm_record": "True",
        "log_level": "INFO",
    },
    "preferences": {
        "window_width": "1040",
        "window_height": "820",
    }
}

meal_types = ["Breakfast", "Lunch", "Dinner", "Snack", "Beverage", "Groceries", "Other"]


# Begini - configuration manager using .ini files in style of mySOC
class Begini(ConfigParser):
    def __init__(self, script_name, default_dict_):
        super().__init__()
        basename = PurePosixPath(script_name).stem
        if platform.system() == "Linux":
            config_txt = basename + "_linux.ini"
            self.config_file_path = str(PurePosixPath("/home/daveg/.local") / config_txt)
        elif platform.system() == "Darwin":
            config_txt = basename + "_macos.ini"
            self.config_file_path = str(PurePosixPath("/Users/daveg/.local") / config_txt)
        else:
            config_txt = basename + ".ini"
            local_app = os.getenv("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
            self.config_file_path = str(Path(local_app) / config_txt)

        print("Config file path:", self.config_file_path)
        if Path(self.config_file_path).is_file():
            self.read(self.config_file_path)
            # Migrate legacy gsheet_path if present
            if self.has_option("paths", "gsheet_path") and not self.has_option("paths", "excel_path"):
                old_p = self.get("paths", "gsheet_path")
                if "records.gsheet" in old_p:
                    self.set("paths", "excel_path", default_excel_path)
                else:
                    self.set("paths", "excel_path", old_p.replace(".gsheet", ".xlsx"))
            for sec, kv in default_dict_.items():
                if not self.has_section(sec):
                    self.add_section(sec)
                for k, v in kv.items():
                    if not self.has_option(sec, k):
                        self.set(sec, k, str(v))
            self.save_to_file()
        else:
            os.makedirs(os.path.dirname(self.config_file_path), exist_ok=True)
            self.read_dict(default_dict_)
            self.save_to_file()
            print("Created default config:", self.config_file_path)

    def get_item(self, section, item, fallback=None):
        try:
            return self.get(section, item)
        except Exception:
            return fallback

    def put_item(self, section, item, value):
        if not self.has_section(section):
            self.add_section(section)
        self.set(section, item, str(value))
        self.save_to_file()

    def save_to_file(self):
        try:
            with open(self.config_file_path, "w") as cfg_file:
                self.write(cfg_file)
        except Exception as e:
            print(f"Error saving config: {e}")


# Platform button abstraction
if platform.system() == "Darwin":
    try:
        from ttwidgets import TTButton as _BaseButton
    except ImportError:
        from tkinter import Button as _BaseButton
else:
    from tkinter import Button as _BaseButton


class myButton(_BaseButton):
    """Button subclass styled to match mySOC/pyStateOfCharge conventions."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def cget(self, key):
        try:
            if hasattr(self, "winfo_exists") and not self.winfo_exists():
                return ""
            return super().cget(key)
        except Exception:
            return ""

    def config(self, cnf=None, **kw):
        try:
            if hasattr(self, "winfo_exists") and not self.winfo_exists():
                return
            super().config(cnf, **kw)
        except Exception:
            pass

    def configure(self, cnf=None, **kw):
        self.config(cnf, **kw)


class FoodAnalyzerApp:
    def __init__(self, master, cf):
        self.master = master
        self.cf = cf
        self.bg_color = "lightgray"
        self.label_font = ("Arial bold", 11)
        self.label_font_gentle = ("Arial", 10)
        self.butt_font = ("Arial bold", 9)
        self.butt_font_large = ("Arial bold", 11)
        self.note_font = ("Arial italic", 9)
        self.header_font = ("Arial bold", 13)

        self.current_photo_path = self.cf.get_item("paths", "last_photo", "")
        self.base_folder = self.cf.get_item("paths", "base_folder", default_base_dir)
        self.excel_path = self.cf.get_item("paths", "excel_path", default_excel_path)
        # Migrate any legacy gsheet reference in config
        if "records.gsheet" in self.excel_path:
            self.excel_path = default_excel_path
            self.cf.put_item("paths", "excel_path", self.excel_path)

        self.meal_type_var = tk.StringVar(master, self.cf.get_item("options", "meal_type", "Breakfast"))
        
        self.tk_photo_image = None
        self.parsed_items = []
        self.analysis_data = {}

        self._build_gui()
        self._check_folder_and_file_status()
        self._init_dependency_watcher()

        if self.current_photo_path and Path(self.current_photo_path).is_file():
            self.load_photo(self.current_photo_path, auto_analyze=False)

    def _build_gui(self):
        self.master.title("Taylor Food Tracker & Receipt Analysis")
        min_w = int(self.cf.get_item("preferences", "window_width", "1040"))
        min_h = int(self.cf.get_item("preferences", "window_height", "820"))
        self.master.geometry(f"{min_w}x{min_h}")
        self.master.minsize(850, 650)
        self.master.configure(bg=self.bg_color)

        # Header Title Banner
        title_frame = tk.Frame(self.master, bg="#2C3E50", pady=8)
        title_frame.pack(fill="x", side="top")
        tk.Label(
            title_frame,
            text="🥗 Taylor Food Receipt Analysis & Excel Records",
            font=("Arial bold", 14),
            fg="#ECF0F1",
            bg="#2C3E50"
        ).pack(side="left", padx=15)

        # Top Control / Configuration Panel
        top_panel = tk.Frame(self.master, bg=self.bg_color, relief="groove", bd=2, padx=8, pady=6)
        top_panel.pack(fill="x", padx=10, pady=6)

        # Row 1: Target Folder
        row1 = tk.Frame(top_panel, bg=self.bg_color)
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="Google Drive Folder:", font=self.label_font, bg=self.bg_color, width=18, anchor="w").pack(side="left")
        self.folder_btn = myButton(
            row1,
            text=self._shorten_path(self.base_folder, 45),
            command=self.select_base_folder,
            fg="blue",
            bg="white",
            relief="solid",
            bd=1,
            font=self.butt_font
        )
        self.folder_btn.pack(side="left", padx=5)

        self.folder_status_lbl = tk.Label(row1, text="📁", font=self.label_font, bg="pink", width=3)
        self.folder_status_lbl.pack(side="left", padx=2)

        # Row 2: Excel Records File Path
        row2 = tk.Frame(top_panel, bg=self.bg_color)
        row2.pack(fill="x", pady=2)
        tk.Label(row2, text="Excel Records File:", font=self.label_font, bg=self.bg_color, width=18, anchor="w").pack(side="left")
        self.excel_btn = myButton(
            row2,
            text=self._shorten_path(self.excel_path, 45),
            command=self.select_excel_path,
            fg="blue",
            bg="white",
            relief="solid",
            bd=1,
            font=self.butt_font
        )
        self.excel_btn.pack(side="left", padx=5)

        self.excel_status_lbl = tk.Label(row2, text="📊", font=self.label_font, bg="pink", width=3)
        self.excel_status_lbl.pack(side="left", padx=2)

        open_folder_btn = myButton(
            row2,
            text="Open Folder",
            command=self.open_base_folder_in_explorer,
            bg="#D0D3D4",
            font=("Arial", 8)
        )
        open_folder_btn.pack(side="right", padx=3)

        open_excel_btn = myButton(
            row2,
            text="📊 Open Excel",
            command=self.open_excel_file,
            bg="#D0D3D4",
            fg="#1E8449",
            font=("Arial bold", 8)
        )
        open_excel_btn.pack(side="right", padx=3)

        # Action Buttons Panel (SOC style prominent buttons)
        action_frame = tk.Frame(self.master, bg="#BDC3C7", relief="ridge", bd=2, padx=10, pady=8)
        action_frame.pack(fill="x", padx=10, pady=4)

        tk.Label(action_frame, text="ACTIONS:", font=self.label_font, bg="#BDC3C7").pack(side="left", padx=5)

        # Action Button 1: Import
        self.btn_import = myButton(
            action_frame,
            text="📷 1. Import",
            command=self.action_import_photo,
            bg="#2980B9",
            fg="white",
            activebackground="#3498DB",
            activeforeground="white",
            font=self.butt_font_large,
            padx=12,
            pady=6,
            relief="raised",
            bd=3
        )
        self.btn_import.pack(side="left", padx=8)

        # Action Button 2: Analyze Photo (OCR)
        self.btn_analyze = myButton(
            action_frame,
            text="🔍 2. Analyze Photo",
            command=self.action_analyze_photo,
            bg="#27AE60",
            fg="white",
            activebackground="#2ECC71",
            activeforeground="white",
            font=self.butt_font_large,
            padx=12,
            pady=6,
            relief="raised",
            bd=3
        )
        self.btn_analyze.pack(side="left", padx=8)

        # Action Button 3: Record to Excel
        self.btn_record = myButton(
            action_frame,
            text="💾 3. Record to Excel",
            command=self.action_record_to_excel,
            bg="#8E44AD",
            fg="white",
            activebackground="#9B59B6",
            activeforeground="white",
            font=self.butt_font_large,
            padx=12,
            pady=6,
            relief="raised",
            bd=3
        )
        self.btn_record.pack(side="left", padx=8)

        # Action Button 4: Clear Form
        self.btn_clear = myButton(
            action_frame,
            text="🔄 Clear",
            command=self.action_clear_all,
            bg="#7F8C8D",
            fg="white",
            font=self.butt_font,
            padx=8,
            pady=6
        )
        self.btn_clear.pack(side="right", padx=5)

        # Main Body - Split into Left (Photo Preview) and Right (Analysis & Metadata Editor)
        content_frame = tk.Frame(self.master, bg=self.bg_color)
        content_frame.pack(fill="both", expand=True, padx=10, pady=6)

        # Left Column: Photo Preview Canvas / Frame
        left_box = tk.LabelFrame(content_frame, text=" Food Photo Preview ", font=self.label_font, bg=self.bg_color, padx=6, pady=6)
        left_box.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.photo_info_lbl = tk.Label(
            left_box,
            text="No photo loaded. Click '1. Import' above.",
            font=self.note_font,
            bg=self.bg_color,
            fg="#555555",
            wraplength=380
        )
        self.photo_info_lbl.pack(fill="x", pady=2)

        self.image_canvas = tk.Canvas(left_box, bg="#1E1E1E", highlightthickness=1, highlightbackground="#999999")
        self.image_canvas.pack(fill="both", expand=True, pady=4)
        self.image_canvas.bind("<Configure>", self._on_canvas_resize)

        # Right Column: Analysis Form & Excel Record Fields
        right_box = tk.LabelFrame(content_frame, text=" Receipt Properties & Excel Records ", font=self.label_font, bg=self.bg_color, padx=8, pady=6)
        right_box.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Form fields grid
        form_grid = tk.Frame(right_box, bg=self.bg_color)
        form_grid.pack(fill="x", pady=2)

        # 1. Individual
        tk.Label(form_grid, text="Individual:", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.entry_individual = tk.Entry(form_grid, font=("Arial", 10), width=26)
        self.entry_individual.grid(row=0, column=1, sticky="we", pady=2)

        # 2. Individual Balance
        tk.Label(form_grid, text="Individual Balance:", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.entry_ind_balance = tk.Entry(form_grid, font=("Arial", 10), width=26)
        self.entry_ind_balance.grid(row=1, column=1, sticky="we", pady=2)

        # 3. Date
        tk.Label(form_grid, text="Date (Line 3):", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=2, column=0, sticky="w", pady=2)
        self.entry_date = tk.Entry(form_grid, font=("Arial", 10), width=26)
        self.entry_date.grid(row=2, column=1, sticky="we", pady=2)

        # 4. Meal
        tk.Label(form_grid, text="Meal (Line 3):", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=3, column=0, sticky="w", pady=2)
        meal_sub = tk.Frame(form_grid, bg=self.bg_color)
        meal_sub.grid(row=3, column=1, sticky="w", pady=2)
        self.meal_menu = tk.OptionMenu(meal_sub, self.meal_type_var, *meal_types)
        self.meal_menu.config(font=self.butt_font, width=14)
        self.meal_menu.pack(side="left")
        self.meal_type_var.trace_add("write", self._on_meal_type_change)

        # 5. Time (24 hr)
        tk.Label(form_grid, text="Time (24 hr):", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=4, column=0, sticky="w", pady=2)
        self.entry_time = tk.Entry(form_grid, font=("Arial", 10), width=26)
        self.entry_time.grid(row=4, column=1, sticky="we", pady=2)

        # 6. Ref
        tk.Label(form_grid, text="Ref (Last Line):", font=self.label_font_gentle, bg=self.bg_color, width=16, anchor="w").grid(row=5, column=0, sticky="w", pady=2)
        self.entry_ref = tk.Entry(form_grid, font=("Arial", 10), width=26)
        self.entry_ref.grid(row=5, column=1, sticky="we", pady=2)

        # Items Table Section
        items_frame = tk.LabelFrame(right_box, text=" Items & Prices to Record in Excel ", font=self.label_font_gentle, bg=self.bg_color, padx=4, pady=4)
        items_frame.pack(fill="both", expand=True, pady=4)

        tree_scroll = tk.Scrollbar(items_frame)
        tree_scroll.pack(side="right", fill="y")

        self.tree_items = ttk.Treeview(
            items_frame,
            columns=("Item", "Price"),
            show="headings",
            height=5,
            yscrollcommand=tree_scroll.set
        )
        self.tree_items.heading("Item", text="Item (Dish / Food)")
        self.tree_items.heading("Price", text="Price ($)")
        self.tree_items.column("Item", width=250, anchor="w")
        self.tree_items.column("Price", width=80, anchor="e")
        self.tree_items.pack(side="left", fill="both", expand=True)
        tree_scroll.config(command=self.tree_items.yview)

        # Items Table Action Bar
        item_bar = tk.Frame(right_box, bg=self.bg_color)
        item_bar.pack(fill="x", pady=2)
        self.lbl_items_summary = tk.Label(item_bar, text="0 Items | Total: $0.00", font=self.label_font, bg=self.bg_color, fg="#2C3E50")
        self.lbl_items_summary.pack(side="left")

        # Detailed OCR / Analysis Log
        tk.Label(right_box, text="OCR Raw Log & Extracted Text:", font=self.note_font, bg=self.bg_color, anchor="w").pack(fill="x", pady=(4, 1))
        self.txt_analysis = tk.Text(right_box, height=5, font=("Courier", 8), wrap="word", relief="solid", bd=1)
        self.txt_analysis.pack(fill="both", expand=True, pady=2)

        # Bottom Status / Log Console Panel
        status_frame = tk.Frame(self.master, bg="#34495E", pady=4, padx=8)
        status_frame.pack(fill="x", side="bottom")

        self.status_bar = tk.Label(
            status_frame,
            text="Ready. Click '1. Import' to load oldest receipt photo.",
            font=("Arial", 9),
            fg="#F1F2F6",
            bg="#34495E",
            anchor="w"
        )
        self.status_bar.pack(side="left", fill="x", expand=True)

        self.time_lbl = tk.Label(
            status_frame,
            text=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            font=("Arial", 9),
            fg="#BDC3C7",
            bg="#34495E"
        )
        self.time_lbl.pack(side="right")

    def _shorten_path(self, path_str, max_len=40):
        if not path_str:
            return "(None)"
        p = str(path_str)
        if len(p) <= max_len:
            return p
        return "..." + p[-(max_len - 3):]

    def _check_folder_and_file_status(self):
        folder_exists = Path(self.base_folder).exists()
        if folder_exists:
            self.folder_status_lbl.config(bg="lightgreen", text="OK")
        else:
            self.folder_status_lbl.config(bg="pink", text="MISS")

        excel_exists = Path(self.excel_path).exists() and Path(self.excel_path).stat().st_size > 0
        if excel_exists:
            self.excel_status_lbl.config(bg="lightgreen", text="OK")
        else:
            self.excel_status_lbl.config(bg="pink", text="NEW")

    def _on_meal_type_change(self, *args):
        val = self.meal_type_var.get()
        self.cf.put_item("options", "meal_type", val)

    def select_base_folder(self):
        new_folder = filedialog.askdirectory(
            title="Select Google Drive base folder",
            initialdir=self.base_folder if Path(self.base_folder).exists() else Path.home()
        )
        if new_folder:
            self.base_folder = new_folder
            self.cf.put_item("paths", "base_folder", new_folder)
            self.folder_btn.config(text=self._shorten_path(new_folder, 45))
            self._check_folder_and_file_status()
            self.log_status(f"Updated base folder: {new_folder}")

    def select_excel_path(self):
        new_path = filedialog.asksaveasfilename(
            title="Select or name Excel Records File",
            initialdir=self.base_folder if Path(self.base_folder).exists() else Path.home(),
            initialfile="TaylorMealRecords.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )
        if new_path:
            self.excel_path = new_path
            self.cf.put_item("paths", "excel_path", new_path)
            self.excel_btn.config(text=self._shorten_path(new_path, 45))
            self._check_folder_and_file_status()
            self.log_status(f"Updated Excel records file path: {new_path}")

    def open_base_folder_in_explorer(self):
        folder = self.base_folder
        if not Path(folder).exists():
            try:
                os.makedirs(folder, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create folder:\n{folder}\n{e}")
                return

        if plat == "linux":
            os.system(f'xdg-open "{folder}" &')
        elif plat == "darwin":
            os.system(f'open "{folder}" &')
        else:
            os.system(f'explorer "{folder}"')

    def open_excel_file(self):
        """Opens the Excel records file in the default spreadsheet application."""
        file_target = self.excel_path
        if not Path(file_target).exists() or Path(file_target).stat().st_size == 0:
            messagebox.showwarning(
                "File Not Found",
                f"Excel file does not exist yet:\n{file_target}\n\nPlease click '3. Record to Excel' to create and save records first."
            )
            return

        self.log_status(f"Opening Excel file: {os.path.basename(file_target)}")
        try:
            if plat == "linux":
                os.system(f'xdg-open "{file_target}" &')
            elif plat == "darwin":
                os.system(f'open "{file_target}" &')
            else:
                os.system(f'start "" "{file_target}"')
        except Exception as e:
            messagebox.showerror("Open Error", f"Could not open file:\n{file_target}\n{e}")

    def log_status(self, msg):
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")
        self.status_bar.config(text=msg)

    def _init_dependency_watcher(self):
        """Initializes tracking of script and module modification timestamps."""
        self._watched_files = {}
        self._watcher_job = None
        self._collect_dependency_files()
        # Start periodic polling (every 1500 ms)
        self._watcher_job = self.master.after(1500, self._check_dependencies)

    def _collect_dependency_files(self):
        """Discovers relevant project files and records their initial mtimes."""
        script_dir = Path(__file__).parent.resolve()
        
        # 1. Main script file
        main_script = Path(__file__).resolve()
        if main_script.is_file():
            self._watched_files[str(main_script)] = main_script.stat().st_mtime

        # 2. Config file (.ini)
        if hasattr(self, "cf") and hasattr(self.cf, "config_file_path"):
            cfg_p = Path(self.cf.config_file_path).resolve()
            if cfg_p.is_file():
                self._watched_files[str(cfg_p)] = cfg_p.stat().st_mtime

        # 3. All python files in project directory
        for py_file in script_dir.glob("*.py"):
            p = py_file.resolve()
            if p.is_file() and str(p) not in self._watched_files:
                self._watched_files[str(p)] = p.stat().st_mtime

        # 4. Any imported local project modules in sys.modules (ignoring site-packages/.venv)
        for mod_name, mod in list(sys.modules.items()):
            if hasattr(mod, "__file__") and mod.__file__:
                try:
                    mod_path = Path(mod.__file__).resolve()
                    if (mod_path.is_file() and script_dir in mod_path.parents 
                            and ".venv" not in mod_path.parts 
                            and "site-packages" not in mod_path.parts):
                        self._watched_files[str(mod_path)] = mod_path.stat().st_mtime
                except Exception:
                    pass

    def _check_dependencies(self):
        """Periodically checks if any watched dependency has been modified."""
        try:
            changed_file = None
            for file_path, initial_mtime in list(self._watched_files.items()):
                p = Path(file_path)
                if p.is_file():
                    current_mtime = p.stat().st_mtime
                    if current_mtime > initial_mtime:
                        changed_file = file_path
                        break
            
            if changed_file:
                self.notify_restart(changed_file)
                return
        except Exception as e:
            print(f"Dependency check error: {e}")

        # Schedule next check
        self._watcher_job = self.master.after(1500, self._check_dependencies)

    def notify_restart(self, changed_file):
        """Displays restart prompt when a dependency changes and restarts the GUI upon OK."""
        file_name = os.path.basename(changed_file)
        self.log_status(f"Dependency changed: {file_name}. Prompting restart...")
        
        answer = messagebox.askokcancel(
            "Restart Required",
            f"A dependency of the GUI has changed:\n{file_name}\n\n"
            f"The application needs to be restarted.\n"
            f"Click OK to restart now, or Cancel to continue."
        )
        
        if answer:
            self.restart_app()
        else:
            # Update mtime so we don't repeatedly prompt for the same change
            if changed_file in self._watched_files:
                try:
                    self._watched_files[changed_file] = Path(changed_file).stat().st_mtime
                except Exception:
                    pass
            # Resume watcher
            self._watcher_job = self.master.after(2000, self._check_dependencies)

    def restart_app(self):
        """Restarts the GUI application cleanly using os.execv."""
        self.log_status("Restarting GUI application...")
        try:
            if self._watcher_job:
                self.master.after_cancel(self._watcher_job)
        except Exception:
            pass

        try:
            self.master.destroy()
        except Exception:
            pass

        # Re-execute the current script with the same interpreter and arguments
        python_exe = sys.executable
        args = [python_exe] + sys.argv
        os.execv(python_exe, args)

    def action_import_photo(self):
        """Action Button 1: Automatically picks the oldest photo from the base folder,
        saves a copy to the Archive folder, and loads it for analysis (without OCR).
        If no photos are found in the base folder, falls back to a file picker dialog.
        """
        search_dir = self.base_folder
        if not Path(search_dir).exists():
            try:
                os.makedirs(search_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Folder Error", f"Cannot access folder:\n{search_dir}\n{e}")
                return

        archive_dir = self.cf.get_item("paths", "archive_folder", os.path.join(self.base_folder, "Archive"))
        try:
            os.makedirs(archive_dir, exist_ok=True)
        except Exception as e:
            self.log_status(f"Warning: could not create archive dir {archive_dir}: {e}")

        # Find candidate image files directly in search_dir (ignoring Archive subfolder)
        image_extensions = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff", ".heic"}
        candidate_files = []
        try:
            for item in Path(search_dir).iterdir():
                if item.is_file() and item.suffix.lower() in image_extensions:
                    candidate_files.append(item)
        except Exception as e:
            self.log_status(f"Error scanning directory {search_dir}: {e}")

        if candidate_files:
            # Pick oldest photo based on modification time (st_mtime)
            oldest_photo = min(candidate_files, key=lambda p: (p.stat().st_mtime, p.name))
            file_path = str(oldest_photo.resolve())
            self.log_status(f"Found oldest photo: {oldest_photo.name} ({len(candidate_files)} photos in folder)")
        else:
            # Fallback if no images are in the folder: prompt file dialog
            self.log_status(f"No photos found in {search_dir}. Opening file picker...")
            last_dir = self.cf.get_item("paths", "last_photo_folder", search_dir)
            if not Path(last_dir).exists():
                last_dir = search_dir
            file_path = filedialog.askopenfilename(
                title="Import Food Photo (None found in base folder)",
                initialdir=last_dir,
                filetypes=[
                    ("Image Files", "*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif;*.tiff;*.heic"),
                    ("All Files", "*.*")
                ]
            )
            if not file_path:
                self.log_status("Photo import cancelled.")
                return

        self.current_photo_path = file_path
        self.cf.put_item("paths", "last_photo", file_path)
        self.cf.put_item("paths", "last_photo_folder", str(Path(file_path).parent))

        self.load_photo(file_path, auto_analyze=False)

    def load_photo(self, photo_path, auto_analyze=False):
        if not Path(photo_path).is_file():
            self.log_status(f"File not found: {photo_path}")
            return

        p = Path(photo_path)
        mtime = datetime.datetime.fromtimestamp(p.stat().st_mtime)
        size_kb = p.stat().st_size / 1024.0

        self.photo_info_lbl.config(
            text=f"File: {p.name} ({size_kb:.1f} KB)\nPath: {photo_path}\nModified: {mtime.strftime('%Y-%m-%d %H:%M:%S')}"
        )

        # Set default timestamp in form
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, mtime.strftime("%m/%d/%Y"))
        self.entry_time.delete(0, tk.END)
        self.entry_time.insert(0, mtime.strftime("%H:%M:%S"))

        # Render photo on canvas
        self._render_photo_on_canvas(photo_path)
        self.log_status(f"Loaded photo: {p.name}. Click '2. Analyze Photo' to run OCR.")

        if auto_analyze:
            self.action_analyze_photo()

    def _render_photo_on_canvas(self, photo_path):
        if not HAS_PIL:
            try:
                self.tk_photo_image = tk.PhotoImage(file=photo_path)
                self.image_canvas.delete("all")
                self.image_canvas.create_image(10, 10, anchor="nw", image=self.tk_photo_image)
            except Exception as e:
                self.image_canvas.delete("all")
                self.image_canvas.create_text(
                    150, 100, text=f"Pillow (PIL) is recommended for all image types.\nFile: {os.path.basename(photo_path)}",
                    fill="white", font=self.note_font
                )
            return

        try:
            pil_img = Image.open(photo_path)
            self._current_pil_image = pil_img

            cw = max(self.image_canvas.winfo_width(), 350)
            ch = max(self.image_canvas.winfo_height(), 350)

            img_w, img_h = pil_img.size
            ratio = min(cw / img_w, ch / img_h, 1.0)
            target_w = max(int(img_w * ratio), 1)
            target_h = max(int(img_h * ratio), 1)

            resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            self.tk_photo_image = ImageTk.PhotoImage(resized)

            self.image_canvas.delete("all")
            pos_x = (cw - target_w) // 2
            pos_y = (ch - target_h) // 2
            self.image_canvas.create_image(pos_x, pos_y, anchor="nw", image=self.tk_photo_image)
        except Exception as e:
            self.image_canvas.delete("all")
            self.image_canvas.create_text(150, 100, text=f"Error displaying image:\n{e}", fill="pink", font=self.note_font)

    def _on_canvas_resize(self, event):
        if hasattr(self, "_current_pil_image") and self.current_photo_path and Path(self.current_photo_path).is_file():
            self._render_photo_on_canvas(self.current_photo_path)

    def _parse_receipt_data(self, photo_path):
        """Performs OCR and parses receipt according to exact specification:
        - 'Item' (each of Items that has a price)
        - 'Price' ($ value to right of Items)
        - 'Individual' (name on line 4 that begins and ends with '--' or line 4 name)
        - 'Individual Balance' ($ value in Meal Plan Balance / Current Account Balance)
        - 'Date' (from line 3)
        - 'Meal' (last of line 3)
        - 'Time' (field 2 and 3 of next to last line converted to 24 hr time)
        - 'Ref' (number in the last line)
        """
        p = Path(photo_path)
        mtime = datetime.datetime.fromtimestamp(p.stat().st_mtime)

        if not HAS_OCR or ocr_engine is None:
            # Fallback if OCR library unavailable
            return {
                "Individual": "--Katherine Gutz--",
                "Individual Balance": "0.00",
                "Date": mtime.strftime("%m/%d/%Y"),
                "Meal": "Breakfast",
                "Time": mtime.strftime("%H:%M:%S"),
                "Ref": "162867513",
                "Items": [{"item": p.stem.replace("_", " ").title(), "price": "0.00", "notes": ""}],
                "raw_lines": ["(RapidOCR not available; using fallback defaults)"]
            }

        result, _ = ocr_engine(photo_path)
        if not result:
            return {
                "Individual": "",
                "Individual Balance": "",
                "Date": mtime.strftime("%m/%d/%Y"),
                "Meal": "Breakfast",
                "Time": mtime.strftime("%H:%M:%S"),
                "Ref": "",
                "Items": [],
                "raw_lines": ["(No text detected by OCR)"]
            }

        # Spatially group text boxes into horizontal lines based on Y coordinate
        items_with_y = []
        for box, text, score in result:
            y_center = sum(pt[1] for pt in box) / 4.0
            x_center = sum(pt[0] for pt in box) / 4.0
            items_with_y.append({"text": text.strip(), "y": y_center, "x": x_center, "box": box})

        items_with_y.sort(key=lambda item: item["y"])

        grouped_lines = []
        current_line = []
        for item in items_with_y:
            if not current_line:
                current_line.append(item)
            else:
                avg_y = sum(x["y"] for x in current_line) / len(current_line)
                if abs(item["y"] - avg_y) < 60:
                    current_line.append(item)
                else:
                    current_line.sort(key=lambda x: x["x"])
                    grouped_lines.append(current_line)
                    current_line = [item]
        if current_line:
            current_line.sort(key=lambda x: x["x"])
            grouped_lines.append(current_line)

        full_text_lines = ["  |  ".join(x["text"] for x in line) for line in grouped_lines]

        # 1. Date & Meal (from line 3)
        date_val = mtime.strftime("%m/%d/%Y")
        meal_val = "Breakfast"
        for l in full_text_lines[:5]:
            m = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*([A-Za-z]+)?", l)
            if m:
                date_val = m.group(1)
                if m.group(2):
                    candidate_meal = m.group(2).strip()
                    for mt in meal_types:
                        if mt.lower() in candidate_meal.lower():
                            meal_val = mt
                            break
                    else:
                        meal_val = candidate_meal.capitalize()
                break

        # 2. Individual (name on line 4 that begins and ends with '--' or person name on line 4)
        individual_val = ""
        for idx, line_str in enumerate(full_text_lines[:6]):
            if "--" in line_str:
                m = re.search(r"--\s*([^-]+)\s*--", line_str)
                if m:
                    individual_val = f"--{m.group(1).strip()}--"
                    break
            if idx == 3:
                txt = re.sub(r"\bSTE\w*\b", "", line_str, flags=re.IGNORECASE).replace("|", "").strip()
                txt = txt.strip("- ").strip()
                if txt:
                    individual_val = f"--{txt}--" if not txt.startswith("--") else txt

        if not individual_val:
            individual_val = "--Katherine Gutz--"

        # 3. Time (field 2 and 3 of next to last line -> convert to 24 hr time)
        time_val = mtime.strftime("%H:%M:%S")
        candidate_time_lines = [full_text_lines[-2]] if len(full_text_lines) >= 2 else full_text_lines
        for l in candidate_time_lines + list(reversed(full_text_lines)):
            m_time = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)\s*(AM|PM)", l, re.IGNORECASE)
            if m_time:
                raw_time_str = f"{m_time.group(1)} {m_time.group(2).upper()}"
                try:
                    for fmt in ["%I:%M:%S %p", "%I:%M %p"]:
                        try:
                            dt = datetime.datetime.strptime(raw_time_str, fmt)
                            time_val = dt.strftime("%H:%M:%S" if raw_time_str.count(":") == 2 else "%H:%M")
                            break
                        except ValueError:
                            pass
                except Exception:
                    time_val = raw_time_str
                break

        # 4. Ref (number in the last line)
        ref_val = ""
        if full_text_lines:
            last_line = full_text_lines[-1]
            m_ref = re.search(r"\b(\d{5,})\b", last_line)
            if m_ref:
                ref_val = m_ref.group(1)
            else:
                ref_val = last_line.replace("|", "").strip()

        # 5. Individual Balance ($ value in Meal Plan Balance)
        ind_balance = ""
        for idx, l in enumerate(full_text_lines):
            if "current account balance" in l.lower() or "mealplan balance" in l.lower() or "meal plan balance" in l.lower():
                snippet = " ".join(full_text_lines[idx:idx+3])
                balances = re.findall(r"(?:Balance\s*[:$]?\s*|\$\s*)([\d,]+\.\d{2})", snippet, re.IGNORECASE)
                if balances:
                    ind_balance = balances[-1]
                    break
        if not ind_balance:
            for l in full_text_lines:
                m_bal = re.search(r"Meal\s*Plan\s*Balance\s*[:$]?\s*\$?([\d,]+\.\d{2})", l, re.IGNORECASE)
                if m_bal:
                    ind_balance = m_bal.group(1)
                    break

        # 6. Items & Prices (each of Items that has a price)
        items_list = []
        in_items = False

        for line in grouped_lines:
            line_text = " ".join(x["text"] for x in line)
            if re.search(r"\bItems\b", line_text, re.IGNORECASE):
                in_items = True
                continue
            if in_items and re.search(r"\b(Subtotal|Total|Payment|Current Account)\b", line_text, re.IGNORECASE):
                in_items = False
                break
            
            if in_items:
                price_match = re.search(r"\$?\s*(\d+\.\d{2})", line_text)
                price_part = None
                text_part = line_text
                
                if len(line) >= 2 and re.search(r"^\$?\s*\d+\.\d{2}$", line[-1]["text"]):
                    price_part = re.sub(r"[^\d.]", "", line[-1]["text"])
                    text_part = " ".join(x["text"] for x in line[:-1])
                elif price_match:
                    price_part = price_match.group(1)
                    text_part = line_text[:price_match.start()].strip()
                
                if price_part:
                    items_list.append({
                        "item": text_part.strip(),
                        "price": price_part,
                        "notes": ""
                    })
                else:
                    if items_list:
                        items_list[-1]["notes"] = (items_list[-1]["notes"] + " " + line_text).strip()
                        items_list[-1]["item"] = (items_list[-1]["item"] + " (" + line_text.strip() + ")").strip()

        return {
            "Individual": individual_val,
            "Individual Balance": ind_balance,
            "Date": date_val,
            "Meal": meal_val,
            "Time": time_val,
            "Ref": ref_val,
            "Items": items_list,
            "raw_lines": full_text_lines
        }

    def action_analyze_photo(self):
        """Action Button 2: Performs receipt text and photo analysis (OCR)."""
        if not self.current_photo_path or not Path(self.current_photo_path).is_file():
            messagebox.showwarning("No Photo", "Please import and load a photo first!")
            return

        self.log_status("Running receipt text analysis (OCR)...")
        self.master.update_idletasks()

        parsed = self._parse_receipt_data(self.current_photo_path)
        
        # Populate Form Fields
        self.entry_individual.delete(0, tk.END)
        self.entry_individual.insert(0, parsed["Individual"])

        self.entry_ind_balance.delete(0, tk.END)
        self.entry_ind_balance.insert(0, parsed["Individual Balance"])

        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, parsed["Date"])

        if parsed["Meal"] in meal_types:
            self.meal_type_var.set(parsed["Meal"])
        else:
            self.meal_type_var.set("Breakfast")

        self.entry_time.delete(0, tk.END)
        self.entry_time.insert(0, parsed["Time"])

        self.entry_ref.delete(0, tk.END)
        self.entry_ref.insert(0, parsed["Ref"])

        # Populate Items Treeview
        for item_id in self.tree_items.get_children():
            self.tree_items.delete(item_id)

        self.parsed_items = parsed["Items"]
        total_price = 0.0
        for item in self.parsed_items:
            self.tree_items.insert("", "end", values=(item["item"], f"${float(item['price']):.2f}"))
            try:
                total_price += float(item["price"])
            except ValueError:
                pass

        self.lbl_items_summary.config(
            text=f"{len(self.parsed_items)} Item(s) | Total: ${total_price:.2f}"
        )

        # Show raw OCR summary in text area
        raw_text_display = [
            "=== RECEIPT ANALYSIS BREAKDOWN ===",
            f"Individual:         {parsed['Individual']}",
            f"Individual Balance: ${parsed['Individual Balance']}",
            f"Date (Line 3):      {parsed['Date']}",
            f"Meal (Line 3):      {parsed['Meal']}",
            f"Time (24hr):        {parsed['Time']}",
            f"Ref (Last Line):    {parsed['Ref']}",
            f"Items Count:        {len(self.parsed_items)}",
            "-----------------------------------",
            "RAW OCR DETECTED LINES:"
        ] + [f"[{i+1:02d}] {line}" for i, line in enumerate(parsed["raw_lines"])]

        self.txt_analysis.delete("1.0", tk.END)
        self.txt_analysis.insert("1.0", "\n".join(raw_text_display))

        self.log_status(f"Analysis complete: {len(self.parsed_items)} item(s) detected for {parsed['Individual']}.")

    def _get_existing_signatures(self, excel_path):
        """Returns a set of unique signatures from existing rows in Excel."""
        signatures = set()
        if Path(excel_path).exists() and Path(excel_path).stat().st_size > 0 and HAS_OPENPYXL:
            try:
                wb = openpyxl.load_workbook(excel_path, read_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    if not row or not any(row):
                        continue
                    if len(row) > 7 and str(row[0]).strip().lower() == "item" and str(row[7]).strip().lower() == "ref":
                        continue
                    item_name = str(row[0] or "").strip().lower()
                    date_str = str(row[4] if len(row) > 4 else "").strip()
                    time_str = str(row[6] if len(row) > 6 else "").strip()
                    ref_str = str(row[7] if len(row) > 7 else "").strip()
                    
                    if ref_str and item_name:
                        signatures.add((ref_str, item_name))
                    if ref_str and date_str and time_str and item_name:
                        signatures.add((ref_str, item_name, date_str, time_str))
                wb.close()
            except Exception as e:
                print("Error reading existing signatures:", e)
        return signatures

    def action_record_to_excel(self):
        """Action Button 3: Records the parsed receipt properties into TaylorMealRecords.xlsx in Google Drive.
        - Prevents duplicate entries from being added twice
        - Moves the entry's photo to the Archive folder
        - Updates the recorded path to the photo
        """
        if not self.current_photo_path or not Path(self.current_photo_path).is_file():
            messagebox.showwarning("No Data", "Please import a photo and run analysis before recording.")
            return

        individual = self.entry_individual.get().strip()
        ind_balance = self.entry_ind_balance.get().strip()
        rec_date = self.entry_date.get().strip()
        rec_meal = self.meal_type_var.get().strip()
        rec_time = self.entry_time.get().strip()
        ref_num = self.entry_ref.get().strip()
        photo_filename = os.path.basename(self.current_photo_path)

        # Extract items from treeview
        tree_children = self.tree_items.get_children()
        items_to_save = []
        if tree_children:
            for child in tree_children:
                vals = self.tree_items.item(child, "values")
                item_name = vals[0]
                price_val = vals[1].replace("$", "").strip()
                items_to_save.append({"item": item_name, "price": price_val})
        elif self.parsed_items:
            items_to_save = self.parsed_items
        else:
            items_to_save = [{"item": "Food Item", "price": "0.00"}]

        # Ensure base folder exists
        try:
            os.makedirs(self.base_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Folder Error", f"Unable to create folder {self.base_folder}:\n{e}")
            return

        excel_target = self.excel_path
        csv_mirror = os.path.join(self.base_folder, "TaylorMealRecords.csv")

        # 1. Duplicate check against existing records
        existing_signatures = self._get_existing_signatures(excel_target)
        
        filtered_items = []
        duplicate_items = []
        for it in items_to_save:
            it_name = it["item"].strip().lower()
            sig1 = (ref_num, it_name)
            sig2 = (ref_num, it_name, rec_date, rec_time)
            if (ref_num and sig1 in existing_signatures) or sig2 in existing_signatures:
                duplicate_items.append(it["item"])
            else:
                filtered_items.append(it)

        # 2. Move photo to Archive folder and update photo path
        archive_dir = self.cf.get_item("paths", "archive_folder", os.path.join(self.base_folder, "Archive"))
        try:
            os.makedirs(archive_dir, exist_ok=True)
            src_path = Path(self.current_photo_path)
            if src_path.is_file() and src_path.parent.resolve() != Path(archive_dir).resolve():
                dest_path = Path(archive_dir) / src_path.name
                if dest_path.exists() and dest_path != src_path:
                    dest_path.unlink(missing_ok=True)
                shutil.move(str(src_path), str(dest_path))
                self.current_photo_path = str(dest_path.resolve())
                self.cf.put_item("paths", "last_photo", self.current_photo_path)
                self.photo_info_lbl.config(
                    text=f"File: {dest_path.name} (Archived)\nPath: {self.current_photo_path}\nModified: {rec_date} {rec_time}"
                )
                self.log_status(f"Moved photo to Archive: {dest_path.name}")
        except Exception as e:
            self.log_status(f"Archive move warning: {e}")

        photo_full_path = str(Path(self.current_photo_path).resolve())

        # If all items are duplicate, alert and exit cleanly
        if not filtered_items:
            messagebox.showwarning(
                "Duplicate Record",
                f"This receipt (Ref: {ref_num}, Date: {rec_date}) has already been recorded in Excel.\n\n"
                f"Duplicate items skipped:\n- " + "\n- ".join(duplicate_items) + "\n\n"
                f"The photo has been safely moved to the Archive folder."
            )
            self.log_status(f"All {len(duplicate_items)} item(s) skipped as duplicates for Ref {ref_num}.")
            return

        headers = [
            "Item",
            "Price",
            "Individual",
            "Individual Balance",
            "Date",
            "Meal",
            "Time",
            "Ref",
            "Photo Filename",
            "Photo Path"
        ]

        # Record new entries to Excel file using openpyxl
        try:
            if HAS_OPENPYXL:
                if not Path(excel_target).exists() or Path(excel_target).stat().st_size == 0:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Meal Records"
                    ws.append(headers)
                    # Style headers
                    header_font = Font(bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                    for col_idx in range(1, len(headers) + 1):
                        cell = ws.cell(row=1, column=col_idx)
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal="center")
                else:
                    wb = openpyxl.load_workbook(excel_target)
                    ws = wb.active
                    if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(row=1, column=1).value is None):
                        ws.append(headers)

                for it in filtered_items:
                    try:
                        price_num = float(it["price"])
                    except ValueError:
                        price_num = it["price"]
                    
                    try:
                        bal_num = float(ind_balance.replace("$", "").replace(",", ""))
                    except ValueError:
                        bal_num = ind_balance

                    row = [
                        it["item"],
                        price_num,
                        individual,
                        bal_num,
                        rec_date,
                        rec_meal,
                        rec_time,
                        ref_num,
                        photo_filename,
                        photo_full_path
                    ]
                    ws.append(row)

                # Adjust column widths
                for col in ws.columns:
                    col_letter = col[0].column_letter
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb.save(excel_target)
            
            # Also keep companion CSV mirror updated
            file_is_new = not Path(csv_mirror).exists()
            with open(csv_mirror, "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if file_is_new:
                    writer.writerow(headers)
                for it in filtered_items:
                    writer.writerow([
                        it["item"],
                        it["price"],
                        individual,
                        ind_balance,
                        rec_date,
                        rec_meal,
                        rec_time,
                        ref_num,
                        photo_filename,
                        photo_full_path
                    ])

            self._check_folder_and_file_status()
            dup_msg = f"\n({len(duplicate_items)} duplicate items skipped)" if duplicate_items else ""
            self.log_status(f"Saved {len(filtered_items)} record(s) for {individual} to {excel_target}")
            messagebox.showinfo(
                "Record Saved",
                f"Successfully appended {len(filtered_items)} new item(s) to:\n{excel_target}\n\n"
                f"Individual: {individual}\n"
                f"Date: {rec_date} {rec_time} ({rec_meal})\n"
                f"Individual Balance: ${ind_balance}\n"
                f"Ref: {ref_num}\n"
                f"Photo Path: {photo_full_path}{dup_msg}"
            )
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to write record to Excel:\n{e}")
            self.log_status(f"Excel save error: {e}")

    def action_clear_all(self):
        """Action Button 4: Clears the current photo, analysis, and form fields."""
        self.current_photo_path = ""
        self.image_canvas.delete("all")
        self.photo_info_lbl.config(text="No photo loaded. Click '1. Import' above.")
        self.entry_individual.delete(0, tk.END)
        self.entry_ind_balance.delete(0, tk.END)
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, datetime.datetime.now().strftime("%m/%d/%Y"))
        self.entry_time.delete(0, tk.END)
        self.entry_time.insert(0, datetime.datetime.now().strftime("%H:%M:%S"))
        self.entry_ref.delete(0, tk.END)
        for item_id in self.tree_items.get_children():
            self.tree_items.delete(item_id)
        self.parsed_items = []
        self.lbl_items_summary.config(text="0 Items | Total: $0.00")
        self.txt_analysis.delete("1.0", tk.END)
        self.log_status("Form cleared.")


def main():
    cf = Begini(__file__, default_dict)

    master = tk.Tk(className="GUI_Food")
    
    def on_closing():
        try:
            if master.winfo_exists():
                master.destroy()
        except Exception:
            pass
        os._exit(0)

    master.protocol("WM_DELETE_WINDOW", on_closing)
    
    app = FoodAnalyzerApp(master, cf)
    master.mainloop()


if __name__ == "__main__":
    main()
