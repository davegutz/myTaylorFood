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
default_gsheet_path = os.path.join(default_base_dir, "records.gsheet")
default_csv_path = os.path.join(default_base_dir, "records.csv")

# Configuration default values
default_dict = {
    "paths": {
        "base_folder": default_base_dir,
        "archive_folder": default_archive_dir,
        "gsheet_path": default_gsheet_path,
        "last_photo_folder": default_base_dir,
        "last_photo": "",
    },
    "options": {
        "meal_type": "Lunch",
        "auto_analyze": "True",
        "confirm_record": "True",
        "log_level": "INFO",
    },
    "preferences": {
        "window_width": "980",
        "window_height": "750",
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
            # Ensure all default sections and keys exist
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
        self.gsheet_path = self.cf.get_item("paths", "gsheet_path", default_gsheet_path)
        self.meal_type_var = tk.StringVar(master, self.cf.get_item("options", "meal_type", "Lunch"))
        
        self.tk_photo_image = None
        self.analysis_data = {}

        self._build_gui()
        self._check_folder_and_file_status()

        if self.current_photo_path and Path(self.current_photo_path).is_file():
            self.load_photo(self.current_photo_path, auto_analyze=False)

    def _build_gui(self):
        self.master.title("Taylor Food Tracker & Photo Analysis")
        min_w = int(self.cf.get_item("preferences", "window_width", "980"))
        min_h = int(self.cf.get_item("preferences", "window_height", "750"))
        self.master.geometry(f"{min_w}x{min_h}")
        self.master.minsize(800, 600)
        self.master.configure(bg=self.bg_color)

        # Header Title Banner
        title_frame = tk.Frame(self.master, bg="#2C3E50", pady=8)
        title_frame.pack(fill="x", side="top")
        tk.Label(
            title_frame,
            text="🥗 Taylor Food Photo Analysis & Records",
            font=("Arial bold", 14),
            fg="#ECF0F1",
            bg="#2C3E50"
        ).pack(side="left", padx=15)

        # Top Control / Configuration Panel
        top_panel = tk.Frame(self.master, bg=self.bg_color, relief="groove", bd=2, padx=8, pady=6)
        top_panel.pack(fill="x", padx=10, pady=6)

        # Row 1: Target Folder & Sheet Path
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

        # Row 2: Sheet File Path
        row2 = tk.Frame(top_panel, bg=self.bg_color)
        row2.pack(fill="x", pady=2)
        tk.Label(row2, text="Google Sheet File:", font=self.label_font, bg=self.bg_color, width=18, anchor="w").pack(side="left")
        self.sheet_btn = myButton(
            row2,
            text=self._shorten_path(self.gsheet_path, 45),
            command=self.select_gsheet_path,
            fg="blue",
            bg="white",
            relief="solid",
            bd=1,
            font=self.butt_font
        )
        self.sheet_btn.pack(side="left", padx=5)

        self.sheet_status_lbl = tk.Label(row2, text="📊", font=self.label_font, bg="pink", width=3)
        self.sheet_status_lbl.pack(side="left", padx=2)

        open_folder_btn = myButton(
            row2,
            text="Open Folder",
            command=self.open_base_folder_in_explorer,
            bg="#D0D3D4",
            font=("Arial", 8)
        )
        open_folder_btn.pack(side="right", padx=5)

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

        # Action Button 2: Analyze Photo
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

        # Action Button 3: Record to Google Sheet
        self.btn_record = myButton(
            action_frame,
            text="💾 3. Record to Sheet",
            command=self.action_record_to_sheet,
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
            text="No photo loaded. Click '1. Import & Load Photo' above.",
            font=self.note_font,
            bg=self.bg_color,
            fg="#555555",
            wraplength=380
        )
        self.photo_info_lbl.pack(fill="x", pady=2)

        self.image_canvas = tk.Canvas(left_box, bg="#1E1E1E", highlightthickness=1, highlightbackground="#999999")
        self.image_canvas.pack(fill="both", expand=True, pady=4)
        self.image_canvas.bind("<Configure>", self._on_canvas_resize)

        # Right Column: Analysis Form & Google Sheet Record Fields
        right_box = tk.LabelFrame(content_frame, text=" Analysis & Record Details ", font=self.label_font, bg=self.bg_color, padx=8, pady=6)
        right_box.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Form fields
        form_grid = tk.Frame(right_box, bg=self.bg_color)
        form_grid.pack(fill="x", pady=4)

        # Meal Type
        tk.Label(form_grid, text="Meal Type:", font=self.label_font_gentle, bg=self.bg_color, width=14, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        self.meal_menu = tk.OptionMenu(form_grid, self.meal_type_var, *meal_types)
        self.meal_menu.config(font=self.butt_font, width=16)
        self.meal_menu.grid(row=0, column=1, sticky="w", pady=3)
        self.meal_type_var.trace_add("write", self._on_meal_type_change)

        # Item / Dish Name
        tk.Label(form_grid, text="Item / Dish:", font=self.label_font_gentle, bg=self.bg_color, width=14, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
        self.entry_item_name = tk.Entry(form_grid, font=("Arial", 10), width=28)
        self.entry_item_name.grid(row=1, column=1, sticky="we", pady=3)

        # Estimated Cost / Price ($)
        tk.Label(form_grid, text="Cost ($):", font=self.label_font_gentle, bg=self.bg_color, width=14, anchor="w").grid(row=2, column=0, sticky="w", pady=3)
        self.entry_cost = tk.Entry(form_grid, font=("Arial", 10), width=28)
        self.entry_cost.grid(row=2, column=1, sticky="we", pady=3)

        # Calories / Nutrition Est.
        tk.Label(form_grid, text="Est. Calories:", font=self.label_font_gentle, bg=self.bg_color, width=14, anchor="w").grid(row=3, column=0, sticky="w", pady=3)
        self.entry_calories = tk.Entry(form_grid, font=("Arial", 10), width=28)
        self.entry_calories.grid(row=3, column=1, sticky="we", pady=3)

        # Date & Time of Meal / Photo
        tk.Label(form_grid, text="Date / Time:", font=self.label_font_gentle, bg=self.bg_color, width=14, anchor="w").grid(row=4, column=0, sticky="w", pady=3)
        self.entry_datetime = tk.Entry(form_grid, font=("Arial", 10), width=28)
        self.entry_datetime.grid(row=4, column=1, sticky="we", pady=3)
        self.entry_datetime.insert(0, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        # Detailed Notes / Ingredients / Analysis Summary
        tk.Label(right_box, text="Analysis Summary & Notes:", font=self.label_font_gentle, bg=self.bg_color, anchor="w").pack(fill="x", pady=(6, 2))
        self.txt_analysis = tk.Text(right_box, height=8, font=("Courier", 9), wrap="word", relief="solid", bd=1)
        self.txt_analysis.pack(fill="both", expand=True, pady=2)

        # Bottom Status / Log Console Panel (mySOC style)
        status_frame = tk.Frame(self.master, bg="#34495E", pady=4, padx=8)
        status_frame.pack(fill="x", side="bottom")

        self.status_bar = tk.Label(
            status_frame,
            text="Ready. Select or import a food photo to begin.",
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
        # Folder check
        folder_exists = Path(self.base_folder).exists()
        if folder_exists:
            self.folder_status_lbl.config(bg="lightgreen", text="OK")
        else:
            self.folder_status_lbl.config(bg="pink", text="MISS")

        # Sheet check
        sheet_exists = Path(self.gsheet_path).exists() or Path(default_csv_path).exists()
        if sheet_exists:
            self.sheet_status_lbl.config(bg="lightgreen", text="OK")
        else:
            self.sheet_status_lbl.config(bg="pink", text="NEW")

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

    def select_gsheet_path(self):
        new_path = filedialog.asksaveasfilename(
            title="Select or name Google Sheet / record file",
            initialdir=self.base_folder if Path(self.base_folder).exists() else Path.home(),
            initialfile="records.gsheet",
            filetypes=[("Google Sheet / CSV", "*.gsheet;*.csv"), ("All Files", "*.*")]
        )
        if new_path:
            self.gsheet_path = new_path
            self.cf.put_item("paths", "gsheet_path", new_path)
            self.sheet_btn.config(text=self._shorten_path(new_path, 45))
            self._check_folder_and_file_status()
            self.log_status(f"Updated records sheet path: {new_path}")

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

    def log_status(self, msg):
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")
        self.status_bar.config(text=msg)

    def action_import_photo(self):
        """Action Button 1: Automatically picks the oldest photo from the base folder,
        saves a copy to the Archive folder, and loads it for analysis.
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

        # Find candidate image files directly in search_dir (excluding subdirectories)
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

        # Save a copy to Archive
        try:
            dest_archive = Path(archive_dir) / Path(file_path).name
            shutil.copy2(file_path, dest_archive)
            self.log_status(f"Archived copy to: {dest_archive.name}")
        except Exception as e:
            self.log_status(f"Archive copy warning: {e}")

        self.current_photo_path = file_path
        self.cf.put_item("paths", "last_photo", file_path)
        self.cf.put_item("paths", "last_photo_folder", str(Path(file_path).parent))

        self.load_photo(file_path, auto_analyze=True)

    def load_photo(self, photo_path, auto_analyze=True):
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
        self.entry_datetime.delete(0, tk.END)
        self.entry_datetime.insert(0, mtime.strftime("%Y-%m-%d %H:%M:%S"))

        # Render photo on canvas
        self._render_photo_on_canvas(photo_path)

        # Infer dish name from filename if sensible
        stem = p.stem.replace("_", " ").replace("-", " ")
        if not any(char.isdigit() for char in stem) and len(stem) > 2:
            self.entry_item_name.delete(0, tk.END)
            self.entry_item_name.insert(0, stem.title())

        self.log_status(f"Loaded photo: {p.name}")

        if auto_analyze:
            self.action_analyze_photo()

    def _render_photo_on_canvas(self, photo_path):
        if not HAS_PIL:
            # Fallback Tkinter PhotoImage for basic GIF/PNG
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

            # Fit into canvas bounds
            cw = max(self.image_canvas.winfo_width(), 350)
            ch = max(self.image_canvas.winfo_height(), 350)

            # Preserve aspect ratio
            img_w, img_h = pil_img.size
            ratio = min(cw / img_w, ch / img_h, 1.0)
            target_w = max(int(img_w * ratio), 1)
            target_h = max(int(img_h * ratio), 1)

            resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            self.tk_photo_image = ImageTk.PhotoImage(resized)

            self.image_canvas.delete("all")
            # Center on canvas
            pos_x = (cw - target_w) // 2
            pos_y = (ch - target_h) // 2
            self.image_canvas.create_image(pos_x, pos_y, anchor="nw", image=self.tk_photo_image)
        except Exception as e:
            self.image_canvas.delete("all")
            self.image_canvas.create_text(150, 100, text=f"Error displaying image:\n{e}", fill="pink", font=self.note_font)

    def _on_canvas_resize(self, event):
        if hasattr(self, "_current_pil_image") and self.current_photo_path and Path(self.current_photo_path).is_file():
            self._render_photo_on_canvas(self.current_photo_path)

    def action_analyze_photo(self):
        """Action Button 2: Performs food and photo analysis on the loaded image."""
        if not self.current_photo_path or not Path(self.current_photo_path).is_file():
            messagebox.showwarning("No Photo", "Please import and load a photo first!")
            return

        p = Path(self.current_photo_path)
        stat = p.stat()
        mtime = datetime.datetime.fromtimestamp(stat.st_mtime)

        # Image properties
        width, height, img_format, mode = 0, 0, "Unknown", "Unknown"
        if HAS_PIL:
            try:
                with Image.open(self.current_photo_path) as im:
                    width, height = im.size
                    img_format = im.format or p.suffix.upper().replace(".", "")
                    mode = im.mode
            except Exception:
                pass

        # Determine meal type by photo capture time if not manually selected
        hour = mtime.hour
        suggested_meal = "Lunch"
        if 5 <= hour < 11:
            suggested_meal = "Breakfast"
        elif 11 <= hour < 16:
            suggested_meal = "Lunch"
        elif 16 <= hour < 22:
            suggested_meal = "Dinner"
        else:
            suggested_meal = "Snack"

        # Update meal type if current was default
        if self.meal_type_var.get() in meal_types:
            # Keep selected or suggest
            current_choice = self.meal_type_var.get()
        else:
            self.meal_type_var.set(suggested_meal)
            current_choice = suggested_meal

        # Build analysis report
        dish_name = self.entry_item_name.get().strip() or p.stem.replace("_", " ").title()
        if not self.entry_item_name.get().strip():
            self.entry_item_name.delete(0, tk.END)
            self.entry_item_name.insert(0, dish_name)

        summary_lines = [
            f"=== FOOD PHOTO ANALYSIS REPORT ===",
            f"Timestamp:       {mtime.strftime('%Y-%m-%d %H:%M:%S')}",
            f"File Name:       {p.name}",
            f"File Size:       {stat.st_size / 1024.0:.1f} KB",
            f"Dimensions:      {width} x {height} ({img_format}, {mode})",
            f"Meal Category:   {current_choice} (Time-inferred: {suggested_meal})",
            f"Item / Dish:     {dish_name}",
            f"Analysis Status: Ready for verification and Sheet export",
            f"----------------------------------------",
            f"Notes: Extracted image metadata and validated path for Google Drive sync."
        ]

        analysis_text = "\n".join(summary_lines)
        self.txt_analysis.delete("1.0", tk.END)
        self.txt_analysis.insert("1.0", analysis_text)

        self.analysis_data = {
            "timestamp": mtime.strftime("%Y-%m-%d %H:%M:%S"),
            "filename": p.name,
            "filepath": str(p.resolve()),
            "filesize_kb": f"{stat.st_size / 1024.0:.1f}",
            "dimensions": f"{width}x{height}",
            "format": img_format,
            "meal_type": current_choice,
            "dish_name": dish_name,
            "cost": self.entry_cost.get().strip(),
            "calories": self.entry_calories.get().strip(),
            "notes": self.txt_analysis.get("1.0", tk.END).strip()
        }

        self.log_status("Photo analysis completed successfully.")

    def action_record_to_sheet(self):
        """Action Button 3: Appends analysis result into Google Sheet / CSV in Google Drive."""
        if not self.current_photo_path or not Path(self.current_photo_path).is_file():
            messagebox.showwarning("No Data", "Please import a photo and run analysis before recording.")
            return

        # Ensure base folder exists
        try:
            os.makedirs(self.base_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Folder Error", f"Unable to create folder {self.base_folder}:\n{e}")
            return

        dish_name = self.entry_item_name.get().strip() or "Food Item"
        cost = self.entry_cost.get().strip()
        calories = self.entry_calories.get().strip()
        meal_type = self.meal_type_var.get()
        rec_datetime = self.entry_datetime.get().strip() or datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        notes = self.txt_analysis.get("1.0", tk.END).strip().replace("\n", " | ")
        photo_filename = os.path.basename(self.current_photo_path)
        photo_full_path = str(Path(self.current_photo_path).resolve())

        # Header definition for Google Sheet / CSV
        headers = [
            "Timestamp",
            "Date",
            "Time",
            "Meal Type",
            "Item / Dish",
            "Cost ($)",
            "Calories",
            "Photo Filename",
            "Photo Path",
            "Analysis Notes"
        ]

        # Parse date and time components
        dt_parts = rec_datetime.split(" ")
        rec_date = dt_parts[0] if len(dt_parts) > 0 else ""
        rec_time = dt_parts[1] if len(dt_parts) > 1 else ""

        row = [
            rec_datetime,
            rec_date,
            rec_time,
            meal_type,
            dish_name,
            cost,
            calories,
            photo_filename,
            photo_full_path,
            notes
        ]

        # Target file: record to CSV in the folder (which Google Sheet connects to / syncs)
        # Also handle .gsheet file pointer or companion records.csv
        csv_target = os.path.join(self.base_folder, "records.csv")
        file_is_new = not Path(csv_target).exists()

        try:
            with open(csv_target, "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if file_is_new:
                    writer.writerow(headers)
                writer.writerow(row)

            # Also create/update .gsheet helper metadata if needed
            gsheet_file = self.gsheet_path
            if not Path(gsheet_file).exists() and gsheet_file.endswith(".gsheet"):
                try:
                    # Write Google Drive link metadata or companion pointer
                    meta = {
                        "name": "records",
                        "type": "google_sheet",
                        "local_csv_mirror": csv_target,
                        "created": datetime.datetime.now().isoformat()
                    }
                    with open(gsheet_file, "w", encoding="utf-8") as gf:
                        json.dump(meta, gf, indent=2)
                except Exception:
                    pass

            self._check_folder_and_file_status()
            self.log_status(f"Saved record for '{dish_name}' to {csv_target}")
            messagebox.showinfo(
                "Record Saved",
                f"Food record successfully appended to:\n{csv_target}\n\nItem: {dish_name}\nMeal: {meal_type}\nCost: ${cost or '0.00'}"
            )
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to write record to file:\n{e}")
            self.log_status(f"Save error: {e}")

    def action_clear_all(self):
        """Action Button 4: Clears the current photo, analysis, and form fields."""
        self.current_photo_path = ""
        self.image_canvas.delete("all")
        self.photo_info_lbl.config(text="No photo loaded. Click '1. Import & Load Photo' above.")
        self.entry_item_name.delete(0, tk.END)
        self.entry_cost.delete(0, tk.END)
        self.entry_calories.delete(0, tk.END)
        self.entry_datetime.delete(0, tk.END)
        self.entry_datetime.insert(0, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
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
