import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
import time
import pyautogui
import subprocess
import logging

from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from PIL import Image, ImageTk
from PIL import ImageDraw, ImageFont
import numpy as np

import sys
import requests
from packaging import version
from datetime import datetime, timedelta

import re
import random
import pandas as pd
import pyperclip
import shutil
from pygame import mixer  # For audio playback
import threading
VERSION = "1.2.0 Beta"


# Add these imports at the top of your main.py
import sys
import requests
from packaging import version
import json
from datetime import datetime, timedelta

class VersionChecker:
    def __init__(self):
        self.version = VERSION  # Your current version constant
        self.last_check_file = "last_update_check.json"
        self.update_url = "https://github.com/marksalhany/Mark-Salhany/blob/da1fa677495d9beaf408389bf70e4f6488771514/ec0buddy.py"
        # Or use any other URL where you'll host version info
        
    def should_check_updates(self):
        """Check if we should check for updates (once per day)"""
        try:
            if os.path.exists(self.last_check_file):
                with open(self.last_check_file, 'r') as f:
                    data = json.load(f)
                    last_check = datetime.fromisoformat(data['last_check'])
                    if datetime.now() - last_check < timedelta(days=1):
                        return False
            return True
        except:
            return True
            
    def save_last_check(self):
        """Save the timestamp of the last update check"""
        try:
            with open(self.last_check_file, 'w') as f:
                json.dump({
                    'last_check': datetime.now().isoformat()
                }, f)
        except:
            pass
            
    def check_for_updates(self):
        """Check for new versions and notify if available"""
        if not self.should_check_updates():
            return
            
        try:
            response = requests.get(self.update_url, timeout=5)
            if response.status_code == 200:
                latest_version = response.json()['tag_name'].lstrip('v')
                if version.parse(latest_version) > version.parse(self.version):
                    messagebox.showinfo(
                        "Update Available",
                        f"A new version ({latest_version}) of EC0 File Buddy is available!\n\n"
                        f"You are currently running version {self.version}\n\n"
                        "Please contact Mark Salhany for the latest version."
                    )
            self.save_last_check()
        except Exception as e:
            # Silently fail for update checks
            pass

class SettingsManager:
    def __init__(self):
        self.settings_file = 'settings.json'
        self.default_settings = {
            'working_tracker_url': 'https://docs.google.com/spreadsheets/d/11PCbwp0OOqClPqS6IfK_QOUdirSFU3QsnVY8-lQbUfk/edit?gid=0#gid=0',
            'mtexec_path': r"C:\Program Files (x86)\MetroCount v506\Programs\MTExec.exe",            
        }
        self.settings = self.load_settings()
        
        
    def load_settings(self):
        """Load settings from JSON file"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    return json.load(f)
            return self.default_settings.copy()
        except Exception as e:
            print(f"Error loading settings: {e}")
            return self.default_settings.copy()

    def save_settings(self):
        """Save settings to JSON file"""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def get_working_tracker_url(self):
        """Get the current working tracker URL"""
        return self.settings.get('working_tracker_url', self.default_settings['working_tracker_url'])

    def set_working_tracker_url(self, url):
        """Set the working tracker URL"""
        self.settings['working_tracker_url'] = url
        self.save_settings()

    def get_mtexec_path(self):
        """Get the MTExec path"""
        return self.settings.get('mtexec_path', self.default_settings['mtexec_path'])

    def set_mtexec_path(self, path):
        """Set the MTExec path"""
        self.settings['mtexec_path'] = path
        self.save_settings()
   

class MTExecPathDialog:
    def __init__(self, parent, settings_manager):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Set MTExec Path")
        self.dialog.geometry("600x150")
        self.settings_manager = settings_manager
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f"+{x}+{y}")

        # Create and pack widgets
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Current path display
        current_frame = ttk.LabelFrame(main_frame, text="Current MTExec Path", padding="5")
        current_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.current_path = ttk.Label(current_frame, 
                                    text=settings_manager.get_mtexec_path(),
                                    wraplength=550)
        self.current_path.pack(fill=tk.X)

        # New path entry
        new_frame = ttk.LabelFrame(main_frame, text="New Path", padding="5")
        new_frame.pack(fill=tk.X)
        
        path_frame = ttk.Frame(new_frame)
        path_frame.pack(fill=tk.X, pady=5)
        
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=70)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = ttk.Button(path_frame, text="Browse", 
                              command=self.browse_path)
        browse_btn.pack(side=tk.LEFT, padx=2)
        
        set_btn = ttk.Button(path_frame, text="Set Path", 
                            command=self.set_path)
        set_btn.pack(side=tk.LEFT)

    def browse_path(self):
        """Browse for MTExec.exe"""
        filename = filedialog.askopenfilename(
            title="Locate MTExec.exe",
            initialdir="C:/Program Files (x86)/MetroCount v506/Programs",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.path_var.set(filename)

    def set_path(self):
        """Set the new MTExec path"""
        new_path = self.path_var.get().strip()
        if new_path:
            if os.path.exists(new_path) and new_path.lower().endswith('mtexec.exe'):
                self.settings_manager.set_mtexec_path(new_path)
                self.current_path.config(text=new_path)
                messagebox.showinfo("Success", "MTExec path updated successfully!")
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Please select a valid MTExec.exe file")
        else:
            messagebox.showerror("Error", "Please enter a path")

class WorkingTrackerDialog:
    def __init__(self, parent, settings_manager):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Set Working Tracker URL")
        self.dialog.geometry("600x150")
        self.settings_manager = settings_manager                
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog on the parent window
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f"+{x}+{y}")

        # Create and pack widgets
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Current URL display
        current_frame = ttk.LabelFrame(main_frame, text="Current Working Tracker URL", padding="5")
        current_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.current_url = ttk.Label(current_frame, 
                                   text=settings_manager.get_working_tracker_url(),
                                   wraplength=550)
        self.current_url.pack(fill=tk.X)

        # New URL entry
        new_frame = ttk.LabelFrame(main_frame, text="New URL", padding="5")
        new_frame.pack(fill=tk.X)
        
        url_frame = ttk.Frame(new_frame)
        url_frame.pack(fill=tk.X, pady=5)
        
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=70)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        tk.Button(url_frame, text="Set URL", 
            command=self.set_url).pack(side=tk.LEFT)

    def set_url(self):
        """Set the new working tracker URL"""
        new_url = self.url_var.get().strip()
        if new_url:
            # Basic URL validation
            if "docs.google.com/spreadsheets" in new_url:
                self.settings_manager.set_working_tracker_url(new_url)
                self.current_url.config(text=new_url)
                messagebox.showinfo("Success", "Working tracker URL updated successfully!")
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Please enter a valid Google Sheets URL")
        else:
            messagebox.showerror("Error", "Please enter a URL")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind('<Enter>', self.show)
        self.widget.bind('<Leave>', self.hide)

    def show(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 30

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(self.tooltip, text=self.text, 
                         background="#ffffff",
                         relief='solid', borderwidth=1,
                         padding=(5,2))
        label.pack()

    def hide(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class ProcessingType:
    """Processing type constants"""
    STANDARD = "standard"
    CITIES_SPEED = "cities_speed"
    BELLEVUE = "bellevue"

class SheetColumns:
    """Google Sheet column names"""
    MAP_NUM = "Map #"
    SITE_NUM = "Site #"
    STREET = "STREET SET"
    REFERENCE = "REFERENCE ST"
    STATUS = "Status"
    NB_TUBE = "NB/EB TUBE"
    SB_TUBE = "SB/WB TUBE"
    SET_DATE = "Date Set"
    PICKUP_DATE = "DATE P/U"
    ONE_WAY = "One Way"
    PROC_NAMING = "PROCESSING NAMING"
    SPEED_LIMIT = "Speed Limit"

class SiteConfigFrame(ttk.Frame):
    def __init__(self, parent, matched_files=None, process_callback=None, study_length=None, process_type=None):
        super().__init__(parent)
        self.matched_files = matched_files or []
        self.process_callback = process_callback
        self.study_length = study_length
        self.process_type = process_type
        self.logger = logging.getLogger(__name__)
        self.disabled_sites = set()
        self.entries = {}
        self.site_vars = {}

        # Configure main frame for resizing
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Define fixed header widths
        self.headers = [
            ("Map #", 10),
            ("Site #", 15),
            ("NB/EB Box", 12),
            ("SB/WB Box", 12),
            ("Processing Name", 45),
            ("One Way", 10),
            ("Study Length", 12),
            ("Date Set", 12),
            ("Begin Date", 14),
            ("End Date", 14),
            ("Speed Limit", 12)
        ]

        # Create main container
        self.container = ttk.Frame(self)
        self.container.pack(fill=tk.BOTH, expand=True)

        # Header frame
        header_frame = ttk.Frame(self.container)
        header_frame.pack(fill=tk.X)

        # Header tooltip dictionary
        self.tooltip_texts = {
            "Map #": "Unique identifier for the site location corresponding with Kapturrit Map",
            "Site #": "Secondary site identifier, often corresponds with client site list",
            "NB/EB Box": "Box number for North-bound or East-bound direction traffic count",
            "SB/WB Box": "Box number for South-bound or West-bound direction traffic count - if twin set or one-way",
            "Processing Name": "Output filename for the processed data, taken from PROCESSING NAMING in Google Sheet",
            "One Way": "Check this box for one-way streets. Required if site only has SB/WB box",
            "Study Length": "Number of days to process, 1-7 days",
            "Date Set": "Date when tubes were installed",
            "Begin Date": "First day of data processing. Click left/right arrows to adjust",
            "End Date": "Data will process UP to this day, matching MTExec format, calculated from Begin Date and Study Length"            

        }

        # Then replace your existing header creation loop with this:
        for col, (header, width) in enumerate(self.headers):
            cell = ttk.Frame(header_frame, relief='raised', borderwidth=1)
            cell.grid(row=0, column=col, sticky='nsew')
            header_label = ttk.Label(cell, text=header, anchor='center',
                                 font=('TkDefaultFont', 9, 'bold'))
            header_label.pack(padx=2, pady=2, expand=True)
            
            # Add tooltip if text exists for this header
            if header in self.tooltip_texts:
                ToolTip(header_label, self.tooltip_texts[header])
            
            header_frame.grid_columnconfigure(col, weight=0, minsize=width*8)

        # Create scrollable content area
        self.canvas = tk.Canvas(self.container, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self.container, orient="horizontal", command=self.canvas.xview)

        # Pack scrollbars and canvas
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # Content frame
        self.content_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor='nw')

        # Configure content frame columns to match headers exactly
        for col, (_, width) in enumerate(self.headers):
            self.content_frame.grid_columnconfigure(col, weight=0, minsize=width*8)

        # Bind events
        self.content_frame.bind('<Configure>', self._on_frame_configure)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def is_site_disabled(self, site_num):
        """Check if a site is disabled due to validation"""
        vars_dict = self.site_vars[site_num]
        
        # Check speed limit if Cities/Speed processing
        if self.process_type and self.process_type.get() == ProcessingType.CITIES_SPEED:
            speed_limit = str(vars_dict['speed_limit'].get()).strip()
            if not speed_limit:
                return True
            
        # Check one-way configuration
        nb_box = str(vars_dict['nb_box'].get()).strip()
        sb_box = str(vars_dict['sb_box'].get()).strip()
        is_one_way = bool(vars_dict['one_way'].get())
        
        return (not nb_box and sb_box and not is_one_way)

    def update_site_state(self, site_num):
        """Update the enabled/disabled state of a site's fields"""
        if site_num not in self.site_vars:
            return
            
        vars_dict = self.site_vars[site_num]
        disabled = self.is_site_disabled(site_num)
        entries = self.entries.get(site_num, {})
        
        if disabled:
            if self.process_type and self.process_type.get() == ProcessingType.CITIES_SPEED:
                speed_limit = str(vars_dict['speed_limit'].get()).strip()
                if not speed_limit:
                    for field, entry in entries.items():
                        if field == 'speed_limit':
                            # Keep speed limit field editable
                            entry.configure(state='normal')
                        else:
                            # Grey out everything else
                            entry.configure(state='disabled')
                    if site_num not in self.disabled_sites:
                        self.disabled_sites.add(site_num)
                        self.master.master.master.update_status(f"Site {site_num} disabled: Missing speed limit")
            else:
                for field, entry in entries.items():
                    if field in ['proc_name', 'one_way', 'begin_date', 'study_length']:
                        entry.configure(state='normal')
                    elif field in ['map', 'site', 'nb_box', 'sb_box', 'date_set', 'end_date']:
                        entry.configure(state='readonly')
                    else:
                        entry.configure(state='disabled')
                if site_num not in self.disabled_sites:
                    self.disabled_sites.add(site_num)
                    self.master.master.master.update_status(f"Site {site_num} disabled: Missing NB/EB box and not marked as one-way")
        else:
            for field, entry in entries.items():
                if field in ['map', 'site', 'nb_box', 'sb_box', 'date_set', 'end_date']:
                    entry.configure(state='readonly')
                else:
                    entry.configure(state='normal')
            if site_num in self.disabled_sites:
                self.disabled_sites.remove(site_num)

    def get_site_data(self):
        """Get all site data as a dictionary"""
        site_data = {}
        try:
            for site_num, vars_dict in self.site_vars.items():
                site_data[site_num] = {
                    'map': vars_dict['map'].get(),
                    'site': vars_dict['site'].get(),
                    'nb_box': vars_dict['nb_box'].get(),
                    'sb_box': vars_dict['sb_box'].get(),
                    'proc_name': vars_dict['proc_name'].get(),
                    'one_way': vars_dict['one_way'].get(),
                    'study_length': vars_dict['study_length'].get(),
                    'date_set': vars_dict['date_set'].get(),
                    'begin_date': vars_dict['begin_date'].get(),
                    'end_date': vars_dict['end_date'].get(),
                    'speed_limit': vars_dict['speed_limit'].get(),
                    'set_date_parsed': vars_dict['set_date_parsed']
                }
        except Exception as e:
            self.logger.error(f"Error getting site data: {e}")
            
        return site_data

    def format_date_with_day(self, date_obj):
        """Format date as 'Day DD'"""
        return date_obj.strftime("%a %d")

    def parse_date(self, date_str, reference_date=None):
        """Parse date including day name format"""
        try:
            date_str = date_str.strip()
            if not date_str:
                return None

            if ' ' in date_str:  # Day name format (e.g., "Wed 16")
                if reference_date:
                    date = datetime.strptime(f"{date_str} {reference_date.month} {reference_date.year}", "%a %d %m %Y")
                else:
                    now = datetime.now()
                    date = datetime.strptime(f"{date_str} {now.month} {now.year}", "%a %d %m %Y")
                return date
            else:
                try:
                    date = datetime.strptime(date_str, "%m/%d/%Y")
                except ValueError:
                    try:
                        date = datetime.strptime(date_str, "%m/%d/%y")
                    except ValueError:
                        return None
                return date
        except Exception as e:
            self.logger.error(f"Date parsing error: {date_str}")
            return None

    def create_cell(self, parent, row, col, var, readonly=False, site_num=None, field_name=None):
        """Create a cell with fixed width matching header"""
        cell_frame = ttk.Frame(parent, relief='raised', borderwidth=1)
        cell_frame.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)

        entry = ttk.Entry(cell_frame, textvariable=var,
                         state='readonly' if readonly else 'normal',  # Changed back to 'readonly'
                         justify='center',
                         width=self.headers[col][1])
        entry.pack(expand=True, fill='both', padx=2, pady=2)
        
        if site_num is not None and field_name is not None:
            if site_num not in self.entries:
                self.entries[site_num] = {}
            self.entries[site_num][field_name] = entry
            
        if field_name == 'speed_limit':
            var.trace('w', lambda *args: self.validate_speed_limit(site_num))
            
        return entry

    def create_date_cell(self, parent, row, col, var, set_all_button=False, field_name=None, site_num=None):
        """Create a date cell with fixed width matching header"""
        cell_frame = ttk.Frame(parent, relief='raised', borderwidth=1)
        cell_frame.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)

        entry = ttk.Entry(cell_frame, textvariable=var,
                         justify='center',
                         width=self.headers[col][1])
        entry.grid(row=0, column=0, columnspan=3 if set_all_button else 2, sticky='ew', padx=2, pady=2)

        if site_num is not None and field_name is not None:
            if site_num not in self.entries:
                self.entries[site_num] = {}
            self.entries[site_num][field_name] = entry

        btn_frame = ttk.Frame(cell_frame)
        btn_frame.grid(row=1, column=0, columnspan=3 if set_all_button else 2)

        left_btn = ttk.Button(btn_frame, text="←", width=2)
        left_btn.pack(side='left', padx=2)
        left_btn.config(command=lambda: self.adjust_date(var, -1, field_name, site_num))

        right_btn = ttk.Button(btn_frame, text="→", width=2)
        right_btn.pack(side='left', padx=2)
        right_btn.config(command=lambda: self.adjust_date(var, 1, field_name, site_num))

        if set_all_button and field_name == 'begin_date':
            ttk.Button(btn_frame, text="Set All", width=6,
                      command=lambda: self.set_all_dates(var.get(), field_name)).pack(side='left', padx=2)

        return entry

    def create_combobox(self, parent, row, col, var, values):
        """Create a combobox cell with fixed width matching header"""
        cell_frame = ttk.Frame(parent, relief='raised', borderwidth=1)
        cell_frame.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)

        combo = ttk.Combobox(cell_frame, textvariable=var,
                            values=values,
                            state='readonly',
                            justify='center',
                            width=self.headers[col][1]-2)
        combo.pack(expand=True, padx=2, pady=2)
        return combo

    def add_site_row(self, row, file_info):
        """Add a row of site information with proper data access"""
        try:
            map_num = file_info['site_info']['Map #']
            site_info = file_info['site_info']
            self.site_vars[map_num] = {}

            # Map number (readonly)
            map_var = tk.StringVar(value=map_num)
            map_cell = self.create_cell(self.content_frame, row, 0, map_var, True, map_num, 'map')
            self.site_vars[map_num]['map'] = map_var

            # Site number (readonly)
            site_num = site_info.get('Site #', '')
            site_var = tk.StringVar(value=site_num)
            site_cell = self.create_cell(self.content_frame, row, 1, site_var, True, map_num, 'site')
            self.site_vars[map_num]['site'] = site_var

            # NB/EB Box (readonly)
            nb_box = str(site_info.get('NB/EB TUBE', ''))
            nb_var = tk.StringVar(value=nb_box)
            nb_cell = self.create_cell(self.content_frame, row, 2, nb_var, True, map_num, 'nb_box')
            self.site_vars[map_num]['nb_box'] = nb_var

            # SB/WB Box (readonly)
            sb_box = str(site_info.get('SB/WB TUBE', ''))
            sb_var = tk.StringVar(value=sb_box)
            sb_cell = self.create_cell(self.content_frame, row, 3, sb_var, True, map_num, 'sb_box')
            self.site_vars[map_num]['sb_box'] = sb_var

            # Processing Name (editable)
            proc_name = str(site_info.get('PROCESSING NAMING', ''))
            proc_var = tk.StringVar(value=proc_name)
            proc_cell = self.create_cell(self.content_frame, row, 4, proc_var, False, map_num, 'proc_name')
            self.site_vars[map_num]['proc_name'] = proc_var
            self.bind_drag_down(proc_cell)

            # One Way (editable)
            one_way_value = site_info.get('One Way', '')
            is_one_way = str(one_way_value).upper().strip() == 'TRUE'
            one_way_var = tk.BooleanVar(value=is_one_way)
            one_way_frame = ttk.Frame(self.content_frame)
            one_way_frame.grid(row=row, column=5, sticky='nsew', padx=1, pady=1)
            check = ttk.Checkbutton(one_way_frame, variable=one_way_var)
            check.pack(expand=True)
            self.site_vars[map_num]['one_way'] = one_way_var

            # Study Length (editable)
            length_var = tk.StringVar(value=self.study_length.get())
            study_cell = self.create_combobox(self.content_frame, row, 6, length_var,
                                            ["1", "2", "3", "4", "5", "6", "7"])
            self.site_vars[map_num]['study_length'] = length_var    

            # Date Set (readonly)
            date_set = site_info.get('Date Set', '')
            date_var = tk.StringVar(value=date_set)
            date_cell = self.create_cell(self.content_frame, row, 7, date_var, True, map_num, 'date_set')
            self.site_vars[map_num]['date_set'] = date_var

            # Parse Date Set
            set_date = self.parse_date(date_set)
            self.site_vars[map_num]['set_date_parsed'] = set_date

            # Begin Date (editable)
            begin_var = tk.StringVar()
            if set_date:
                begin_date = set_date + timedelta(days=1)
                begin_var.set(self.format_date_with_day(begin_date))
            begin_cell = self.create_date_cell(self.content_frame, row, 8, begin_var,
                                             set_all_button=(row == 0),
                                             field_name='begin_date',
                                             site_num=map_num)
            self.site_vars[map_num]['begin_date'] = begin_var
            self.bind_drag_down(begin_cell)

            # End Date (readonly)
            end_var = tk.StringVar()
            if set_date:
                study_days = int(length_var.get())
                begin_date = self.parse_date(begin_var.get(), reference_date=set_date)
                if begin_date:
                    end_date = begin_date + timedelta(days=study_days)
                    end_var.set(self.format_date_with_day(end_date))
            end_cell = self.create_cell(self.content_frame, row, 9, end_var, readonly=True, site_num=map_num, field_name='end_date')
            self.site_vars[map_num]['end_date'] = end_var

            # Speed Limit (editable)
            speed_limit = str(site_info.get('Speed Limit', ''))
            speed_var = tk.StringVar(value=speed_limit)
            speed_cell = self.create_cell(self.content_frame, row, 10, speed_var, False, map_num, 'speed_limit')
            self.site_vars[map_num]['speed_limit'] = speed_var
            self.bind_drag_down(speed_cell)

            # Update site state and bind change handlers
            self.update_site_state(map_num)
            one_way_var.trace('w', lambda *args: self.update_site_state(map_num))
            speed_var.trace('w', lambda *args: self.update_site_state(map_num))

            # Update end date when values change
            length_var.trace('w', lambda *args: self._update_end_date(map_num))
            begin_var.trace('w', lambda *args: self._update_end_date(map_num))

        except Exception as e:
            self.logger.error(f"Error adding row for Map # {map_num}: {str(e)}")

    def validate_speed_limit(self, site_num):
        """Validate speed limit and update site state"""
        if self.process_type and self.process_type.get() == ProcessingType.CITIES_SPEED:
            vars_dict = self.site_vars[site_num]
            speed_limit = str(vars_dict['speed_limit'].get()).strip()
            self.update_site_state(site_num)

    def adjust_date(self, var, days, field_name, site_num):
        """Adjust date by given number of days"""
        try:
            current_date_str = var.get()
            if not current_date_str:
                return

            set_date = self.site_vars[site_num]['set_date_parsed']
            if not set_date:
                return

            current_date = self.parse_date(current_date_str, reference_date=set_date)
            if not current_date:
                return

            new_date = current_date + timedelta(days=days)

            # Prevent Begin Date from being earlier than Date Set + 1 day
            if field_name == 'begin_date' and new_date < set_date + timedelta(days=1):
                return

            var.set(self.format_date_with_day(new_date))

            # Update End Date if Begin Date changes
            if field_name == 'begin_date':
                self._update_end_date(site_num)

        except Exception as e:
            self.logger.error(f"Error adjusting date: {e}")

    def set_all_dates(self, value, field_name):
        """Set all dates in the specified field to the given value"""
        for site_num, vars_dict in self.site_vars.items():
            if vars_dict[field_name].get() != value:
                vars_dict[field_name].set(value)
                if field_name == 'begin_date':
                    self._update_end_date(site_num)

    def _update_end_date(self, site_num):
        """Update end date based on begin date and study length"""
        try:
            begin_str = self.site_vars[site_num]['begin_date'].get()
            length_str = self.site_vars[site_num]['study_length'].get()

            if begin_str and length_str:
                set_date = self.site_vars[site_num]['set_date_parsed']
                begin_date = self.parse_date(begin_str, reference_date=set_date)
                if begin_date:
                    study_days = int(length_str)
                    end_date = begin_date + timedelta(days=study_days)
                    self.site_vars[site_num]['end_date'].set(
                        self.format_date_with_day(end_date))
        except Exception as e:
            self.logger.error(f"Error updating end date: {e}")

    def _on_frame_configure(self, event):
        """Update scroll region when content frame size changes"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Update canvas window size when canvas size changes"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def bind_drag_down(self, widget):
        """Bind drag down event for quick data entry"""
        widget.bind('<Control-d>', self._drag_down)

    def _drag_down(self, event):
        """Handle drag down event to copy cell value to below cells"""
        widget = event.widget
        value = widget.get()
        info = widget.grid_info()
        col = info['column']
        row = info['row']

        for r in range(row + 1, len(self.matched_files)):
            child = self.content_frame.grid_slaves(row=r, column=col)
            if child:
                child[0].delete(0, 'end')
                child[0].insert(0, value)

    def populate_files(self, matched_files):
        """Populate frame with matched files"""
        try:
            # Clear existing content
            for widget in self.content_frame.winfo_children():
                widget.destroy()

            self.site_vars = {}
            self.entries = {}  # Clear entries dictionary
            self.disabled_sites = set()  # Clear disabled sites set

            if not matched_files:
                self.logger.error("No matched files provided")
                return

            self.matched_files = matched_files.copy()

            # Sort matched files by Map # numerically
            self.matched_files.sort(key=lambda x: self.sort_map_numbers(x['site']))

            # Add each matched file as a row
            row_counter = 0
            for file_info in self.matched_files:
                try:
                    if file_info['site_info'].get('Status') == 'NEEDS TXT FILE' and \
                       file_info['site_info'].get('Map #', '').strip():
                        self.add_site_row(row_counter, file_info)
                        row_counter += 1
                except Exception as e:
                    self.logger.error(f"Error adding row for site {file_info.get('site', 'unknown')}: {e}")
                    continue

        except Exception as e:
            self.logger.error(f"Error populating files: {e}")

    def sort_map_numbers(self, map_num):
        """Convert map number to sortable integer"""
        try:
            num = ''.join(filter(str.isdigit, str(map_num)))
            return int(num) if num else float('inf')
        except:
            return float('inf') 

class TrafficProcessor:
    # Update these constants in your TrafficProcessor class:
    BEGIN_DATE_OFFSET_X = -275
    BEGIN_DATE_OFFSET_Y = -100
    END_DATE_OFFSET_X = 250
    END_DATE_FROM_LEFT_ARROW = 290
    END_DATE_FROM_RIGHT_ARROW = 250
    ARROW_VERTICAL_OFFSET = 28
    ARROW_HORIZONTAL_OFFSET = 4  # Reduced from 12
    ARROW_FROM_TEXT_VERTICAL = 25
    ARROW_FROM_TEXT_HORIZONTAL = 4  # Reduced from 12
    
    def __init__(self):
                
        self.colors = {
            'primary': '#1a73e8',    # Google blue
            'success': '#0f9d58',    # Green
            'warning': '#f4b400',    # Yellow
            'error': '#db4437',      # Red
            'gray': '#5f6368',       # Subtle gray
            'light_gray': '#f1f3f4'  # Background gray
        }
                
        self.check_credentials()
        self.settings_manager = SettingsManager()                
        self.project_data = None
        self.ec0_files = []
        self.mtexec_path = self.settings_manager.get_mtexec_path()       
        mixer.init()  # Initialize the mixer for audio playback
        pyautogui.PAUSE = 0.01                                   
        
        self.setup_logging()
        self.setup_gui()
        self.folder_path.trace('w', self.update_project_name)
        
        # Configure text widget tags for formatting
        self.status.tag_configure('warning', foreground='red', font=('TkDefaultFont', 9, 'bold'))
        self.status.tag_configure('site_title', font=('TkDefaultFont', 9, 'bold'))
        self.status.tag_configure('success', foreground='green')
        self.status.tag_configure('success_header', foreground='green', font=('TkDefaultFont', 11, 'bold'))
        self.status.tag_configure('info', foreground='blue')
        self.status.tag_configure('processing', foreground='purple')
        self.status.tag_configure('bold', font=('TkDefaultFont', 9, 'bold'))
        self.status.tag_configure('header', foreground='blue', font=('TkDefaultFont', 11, 'bold'))
        self.status.tag_configure('warning_header', foreground='red', font=('TkDefaultFont', 11, 'bold'))
        self.status.tag_configure('partial_header', foreground='orange', font=('TkDefaultFont', 11, 'bold'))
                
        
       # Get system DPI scaling factor
        import ctypes
        try:
            awareness = ctypes.c_int()
            ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
            scaling = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
            self.SCALING_FACTOR = scaling
            self.update_status(f"Display scaling factor: {scaling}")
        except:
            self.SCALING_FACTOR = 1.0
            self.update_status("Could not detect display scaling, using default")

        # Start music playback on application start
        # self.is_playing = False
        # self.play_music()

        # Initialize control variables for processing
        self.processing_thread = None
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
       
        if not os.path.exists(self.mtexec_path):
            self.update_status("Warning: MTExec.exe not found at configured path. " + 
                              "Please check Settings > Set MTExec Path", tag='warning')
   
    def check_credentials(self):
        """Check for required credentials and configuration"""
        if not os.path.exists('client_secrets.json'):
            messagebox.showerror(
                "Configuration Error",
                "client_secrets.json not found!\n\n"
                "Please place the client_secrets.json file in the same directory as the executable.\n\n"
                "Contact mark.salhany@idaxdata.com for assistance."
            )
            sys.exit(1)
            
            # Check for settings file on first run
            if not os.path.exists('settings.json'):
                self.show_first_run_dialog()
                
    def show_first_run_dialog(self):
        """Show first-run configuration dialog"""
        dialog = tk.Toplevel()
        dialog.title("First Time Setup")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Welcome message
        ttk.Label(
            dialog,
            text="Welcome to EC0 File Buddy!",
            font=('TkDefaultFont', 12, 'bold')
        ).pack(pady=20)
        
        ttk.Label(
            dialog,
            text="Let's set up a few things before we get started.",
            wraplength=500
        ).pack(pady=10)
        
        # MTExec path frame
        mtexec_frame = ttk.LabelFrame(dialog, text="MTExec Location", padding=10)
        mtexec_frame.pack(fill='x', padx=20, pady=10)
        
        self.mtexec_var = tk.StringVar(
            value=r"C:\Program Files (x86)\MetroCount v506\Programs\MTExec.exe"
        )
        
        ttk.Entry(
            mtexec_frame,
            textvariable=self.mtexec_var,
            width=60
        ).pack(side='left', padx=5)
        
        ttk.Button(
            mtexec_frame,
            text="Browse",
            command=lambda: self.browse_mtexec(self.mtexec_var)
        ).pack(side='left')
        
        # Working tracker frame
        tracker_frame = ttk.LabelFrame(dialog, text="Working Tracker URL", padding=10)
        tracker_frame.pack(fill='x', padx=20, pady=10)
        
        self.tracker_var = tk.StringVar()
        
        ttk.Entry(
            tracker_frame,
            textvariable=self.tracker_var,
            width=70
        ).pack(fill='x')
        
        # Save button
        ttk.Button(
            dialog,
            text="Save Configuration",
            command=lambda: self.save_first_run_config(dialog)
        ).pack(pady=20)
        
    def browse_mtexec(self, var):
        """Browse for MTExec.exe"""
        filename = filedialog.askopenfilename(
            title="Locate MTExec.exe",
            initialdir="C:/Program Files (x86)/MetroCount v506/Programs",
            filetypes=[("Executable files", "*.exe")]
        )
        if filename:
            var.set(filename)
            
    def save_first_run_config(self, dialog):
        """Save first-run configuration"""
        mtexec_path = self.mtexec_var.get().strip()
        tracker_url = self.tracker_var.get().strip()
        
        if not mtexec_path or not os.path.exists(mtexec_path):
            messagebox.showerror(
                "Error",
                "Please select a valid MTExec.exe location"
            )
            return
            
        if not tracker_url:
            messagebox.showerror(
                "Error",
                "Please enter the Working Tracker URL"
            )
            return
            
        # Save settings
        settings = {
            'mtexec_path': mtexec_path,
            'working_tracker_url': tracker_url
        }
        
        with open('settings.json', 'w') as f:
            json.dump(settings, f, indent=4)
            
        dialog.destroy()
        messagebox.showinfo(
            "Setup Complete",
            "Initial configuration completed successfully!\n\n"
            "EC0 File Buddy is ready to use."
        )
        
    def run(self):
        """Start the application"""
        try:
            # Check for updates when app starts
            self.version_checker.check_for_updates()
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"Application error: {str(e)}")
            messagebox.showerror("Error", f"Application error: {str(e)}")
        finally:
            # Your existing cleanup code...
            pass            
                
 
    def find_similar_box_numbers(self, target_num, pickup_folder_path, matched_boxes):
        """Find similar box numbers in EC0 files within a pickup folder"""
        try:
            target = str(target_num)
            similar_matches = []
            
            # Get all EC0 files in the folder
            ec0_files = [f for f in os.listdir(pickup_folder_path) if f.upper().endswith('.EC0')]
            
            for file in ec0_files:
                box_num = self.extract_box_number(file)
                if not box_num:
                    continue
                    
                # Skip if it's an exact match or already matched
                if box_num == target or box_num in matched_boxes:
                    continue
                
                # Check for similar numbers using strict criteria
                if self.is_similar_number(target, box_num):
                    similar_matches.append({
                        'box': box_num,
                        'file': file
                    })
            
            return similar_matches
            
        except Exception as e:
            self.logger.error(f"Error finding similar box numbers: {e}")
            return []

    def update_project_name(self, *args):
        folder = self.folder_path.get()
        if folder:
            self.project_name.set(os.path.basename(folder.rstrip('/\\')))
        else:
            self.project_name.set("No project selected")

    def is_similar_number(self, num1, num2):
        """Check if two numbers are similar based on stricter criteria"""
        try:
            str1, str2 = str(num1), str(num2)
            
            # Must be same length
            if len(str1) != len(str2):
                return False
                
            # Check for one-digit difference
            differences = sum(1 for a, b in zip(str1, str2) if a != b)
            if differences == 1:
                return True
                
            # Check for transposed adjacent digits
            for i in range(len(str1)-1):
                if str1[i] == str2[i+1] and str1[i+1] == str2[i]:
                    return True
                    
            return False
            
        except Exception:
            return False

    def calculate_similarity(self, num1, num2):
        """Calculate a similarity score between two numbers"""
        try:
            str1, str2 = str(num1), str(num2)
            score = 0
            
            # Exact digit matches
            for c1, c2 in zip(str1, str2):
                if c1 == c2:
                    score += 2
                    
            # Numerical proximity if both are numbers
            try:
                if str1.isdigit() and str2.isdigit():
                    diff = abs(int(str1) - int(str2))
                    if diff <= 10:
                        score += (10 - diff)
            except ValueError:
                pass
                
            return score
            
        except Exception:
            return 0

    def create_app_icon(self):
        try:
            icon_path = os.path.join("images", "branding", "ec0_icon.png")
            if not os.path.exists(os.path.dirname(icon_path)):
                os.makedirs(os.path.dirname(icon_path))
                
            # Create a 48x48 icon with higher resolution
            size = 48
            img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
            draw = ImageDraw.Draw(img)
            
            # Modern abstract tube design
            # Main tube shape
            tube_color = '#1a73e8'  # Professional blue
            accent_color = '#4285f4'  # Lighter blue for depth
            
            # Draw stylized tube/data flow
            points = [
                (10, 24),  # Left start
                (20, 14),  # Top curve
                (28, 14),
                (38, 24),  # Right curve
                (28, 34),
                (20, 34),
                (10, 24)   # Back to start
            ]
            
            # Draw main shape
            draw.polygon(points, fill=tube_color)
            
            # Add dynamic lines suggesting data flow
            draw.line([(15, 24), (33, 24)], fill='white', width=2)
            draw.line([(18, 20), (30, 20)], fill='white', width=1)
            draw.line([(18, 28), (30, 28)], fill='white', width=1)
            
            # Add subtle gradient effect
            for i in range(3):
                offset = i * 2
                draw.ellipse([8+offset, 22+offset, 12+offset, 26+offset], 
                            fill=accent_color)
            
            img.save(icon_path, 'PNG')
            return icon_path
        except Exception as e:
            self.logger.error(f"Error creating app icon: {e}")
            return None  

    def show_creature(self):
        """Show a random creature image from the creatures folder"""
        try:
            # Get all image files from the creatures folder
            creatures_path = os.path.join("images", "creatures")
            valid_extensions = ('.png', '.jpg', '.jpeg', '.gif')
            image_files = [f for f in os.listdir(creatures_path) 
                          if f.lower().endswith(valid_extensions)]
            
            if not image_files:
                return
                
            # Select random image
            image_file = random.choice(image_files)
            
            # Create window
            creature_window = tk.Toplevel(self.root)
            creature_window.title("Creature")
            
            # Load and resize image
            img = Image.open(os.path.join(creatures_path, image_file))
            
            # Resize while maintaining aspect ratio
            basewidth = 500
            wpercent = (basewidth / float(img.size[0]))
            hsize = int((float(img.size[1]) * float(wpercent)))
            img = img.resize((basewidth, hsize), Image.Resampling.LANCZOS)
            
            photo = ImageTk.PhotoImage(img)
            
            # Display image
            label = ttk.Label(creature_window, image=photo)
            label.image = photo  # Keep a reference
            label.pack(padx=10, pady=10)
            
            # Center window on screen
            creature_window.update_idletasks()
            width = creature_window.winfo_width()
            height = creature_window.winfo_height()
            x = (creature_window.winfo_screenwidth() // 2) - (width // 2)
            y = (creature_window.winfo_screenheight() // 2) - (height // 2)
            creature_window.geometry(f'+{x}+{y}')

        except Exception as e:
            self.logger.error(f"Error showing creature: {e}")

    def scale_movement(self, value):
        """Scale a movement value based on system DPI scaling"""
        try:
            # Get the scaling factor (should be set during __init__)
            scaling = getattr(self, 'SCALING_FACTOR', 1.0)
            # Apply scaling and round to nearest integer
            scaled_value = int(round(value * scaling))
            return scaled_value
        except Exception as e:
            self.logger.error(f"Error in scale_movement: {e}")
            return value  # Return original value if scaling fails
            
    def setup_logging(self):
        """Initialize logging configuration"""
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"traffic_processor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

    def setup_gui(self):      
        """Initialize all GUI components"""
        self.root = tk.Tk()
        self.root.geometry("1920x900")
        self.root.title(f"Ec0 File Buddy v{VERSION}")
        # Initialize StringVars after root is created
        self.sheet_name = tk.StringVar(value="No tracker loaded")
        self.project_name = tk.StringVar(value="No project selected")
        self.process_type = tk.StringVar(value=ProcessingType.STANDARD)
        self.process_type.trace('w', self.on_process_type_change)

        # Configure main window to be fully responsive
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        # Create menubar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Create Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)

        # Add menu items
        settings_menu.add_command(
            label="Set Working Tracker",
            command=lambda: WorkingTrackerDialog(self.root, self.settings_manager)
        )

        settings_menu.add_command(
            label="Set MTExec Path",
            command=lambda: MTExecPathDialog(self.root, self.settings_manager)
        )

        # Add separator and About option
        settings_menu.add_separator()
        settings_menu.add_command(
            label="About",
            command=lambda: messagebox.showinfo(
                "About Ec0 File Buddy",
                f"Version: {VERSION}\n\n" +
                "Created by Mark Salhany\n" +
                "For IDAX Data Solutions\n\n" +
                "For support: mark.salhany@idaxdata.com\n\n"
                "Have a great day!"
            )
        )

        # Configure styles for headers
        style = ttk.Style()
        style.configure('Header.TLabelframe.Label', font=('TkDefaultFont', 11, 'bold'))
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0')
        style.configure('TEntry', background='#ffffff')
        style.configure('Info.TLabel', font=('TkDefaultFont', 10, 'bold'))

        # Top container for logo and controls
        top_container = ttk.Frame(self.root)
        top_container.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

        # Create left side of top container
        left_frame = ttk.Frame(top_container, width=200)
        left_frame.grid(row=0, column=0, sticky='ns', padx=(5, 20))
        left_frame.grid_propagate(False)  # Prevent frame from shrinking

        # Create right side of top container
        right_frame = ttk.Frame(top_container)
        right_frame.grid(row=0, column=1, sticky='nsew', padx=5)
        top_container.grid_columnconfigure(1, weight=1)

        # Add IDAX logo to left frame
        try:
            logo_img = Image.open("idax.png")
            logo_img = logo_img.resize((150, 75), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = ttk.Label(left_frame, image=self.logo)
            logo_label.pack(anchor='nw', padx=10, pady=5)
        except Exception as e:
            self.logger.error(f"Could not load logo: {e}")

        # Add EC0 icon under IDAX logo
        try:
            ec0_icon_path = os.path.join("images", "branding", "ec0_icon.png")
            if os.path.exists(ec0_icon_path):
                ec0_icon_img = Image.open(ec0_icon_path)
                ec0_icon_img = ec0_icon_img.resize((125, 125), Image.Resampling.LANCZOS)
                self.ec0_icon_photo = ImageTk.PhotoImage(ec0_icon_img)
                ec0_icon_label = ttk.Label(left_frame, image=self.ec0_icon_photo)
                ec0_icon_label.pack(anchor='nw', padx=10, pady=0)
        except Exception as e:
            self.logger.error(f"Could not load EC0 icon: {e}")

        # Project Info section
        project_frame = ttk.LabelFrame(right_frame, text="Project Information", 
                                     padding="5", style='Header.TLabelframe')
        project_frame.pack(fill=tk.X, padx=1, pady=5)

        # URL row
        url_frame = ttk.Frame(project_frame)
        url_frame.pack(fill=tk.X, pady=2)
        url_label = ttk.Label(url_frame, text="Project Tracker URL:")
        url_label.pack(side=tk.LEFT)
        self.sheet_url = tk.StringVar()
        url_entry = ttk.Entry(url_frame, textvariable=self.sheet_url, width=90)
        url_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ToolTip(url_entry, "Enter Google Sheet URL for project tracking data")

        # Buttons frame
        buttons_frame = ttk.Frame(url_frame)
        buttons_frame.pack(side=tk.LEFT, padx=5)
        load_sheet = ttk.Button(buttons_frame, text="Load Sheet", 
                               command=self.load_sheet_data)
        load_sheet.pack(side=tk.LEFT, padx=2)
        ToolTip(load_sheet, "Load data from entered sheet URL")
        
        self.load_tracker_button = ttk.Button(buttons_frame, text="Load Working Tracker",
                                            command=self.load_working_tracker)
        self.load_tracker_button.pack(side=tk.LEFT, padx=2)
        ToolTip(self.load_tracker_button, "Quick load the standard IDAX Working Tracker")

        # Folder row
        folder_frame = ttk.Frame(project_frame)
        folder_frame.pack(fill=tk.X, pady=2)
        folder_label = ttk.Label(folder_frame, text="Project Folder:     ")
        folder_label.pack(side=tk.LEFT)
        ToolTip(folder_label, "Path to project folder containing EC0 files")
        
        self.folder_path = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, width=90)
        folder_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        browse_button = ttk.Button(folder_frame, text="Browse", 
                                 command=self.browse_folder)
        browse_button.pack(side=tk.LEFT)
        ToolTip(browse_button, "Select project folder containing EC0 files")

        # Processing Configuration section
        config_frame = ttk.LabelFrame(right_frame, text="Processing Configuration", 
                                    padding="5", style='Header.TLabelframe')
        config_frame.pack(fill=tk.X, padx=1, pady=5)
       
        # Current Loaded Information section
        current_info_frame = ttk.LabelFrame(right_frame, text="Current Loaded Information", 
                                            padding="5", style='Header.TLabelframe')
        current_info_frame.pack(fill=tk.X, padx=1, pady=5)

        # Sheet Name Display
        sheet_name_frame = ttk.Frame(current_info_frame)
        sheet_name_frame.pack(fill=tk.X, pady=2)
        sheet_name_label = ttk.Label(sheet_name_frame, text="Current Loaded Tracker:")
        sheet_name_label.pack(side=tk.LEFT)
        self.sheet_name_display = ttk.Label(sheet_name_frame, textvariable=self.sheet_name, font=('TkDefaultFont', 10, 'bold'))
        self.sheet_name_display.pack(side=tk.LEFT, padx=5)

        # Project Name Display
        project_name_frame = ttk.Frame(current_info_frame)
        project_name_frame.pack(fill=tk.X, pady=2)
        project_name_label = ttk.Label(project_name_frame, text="Current Loaded Project:")
        project_name_label.pack(side=tk.LEFT)
        self.project_name_display = ttk.Label(project_name_frame, textvariable=self.project_name, font=('TkDefaultFont', 10, 'bold'))
        self.project_name_display.pack(side=tk.LEFT, padx=5)


        # Processing Type
        type_frame = ttk.Frame(config_frame)
        type_frame.pack(fill=tk.X, pady=2)
        ttk.Label(type_frame, text="Processing Type:").pack(side=tk.LEFT)                        

        standard_radio = ttk.Radiobutton(type_frame, text="Standard", 
                                        variable=self.process_type,
                                        value=ProcessingType.STANDARD)
        standard_radio.pack(side=tk.LEFT, padx=10)
        ToolTip(standard_radio, "Basic tube data processing using 15M_ALL STATS profile")

        cities_radio = ttk.Radiobutton(type_frame, text="Cities/Speed", 
                                      variable=self.process_type,
                                      value=ProcessingType.CITIES_SPEED)
        cities_radio.pack(side=tk.LEFT, padx=10)
        ToolTip(cities_radio, "Includes speed info compared against posted speed limits. Uses 15M_ALL STATS_SPEED STATS profile")

        bellevue_radio = ttk.Radiobutton(type_frame, text="Bellevue", 
                                        variable=self.process_type,
                                        value=ProcessingType.BELLEVUE)
        bellevue_radio.pack(side=tk.LEFT, padx=10)
        ToolTip(bellevue_radio, "Produces two .txt files: 7-day processing plus 7-day with Friday to Monday masked out")

        # Study Configuration
        study_frame = ttk.Frame(config_frame)
        study_frame.pack(fill=tk.X, pady=2)

        # Study Length
        length_frame = ttk.Frame(study_frame)
        length_frame.pack(side=tk.LEFT)
        ttk.Label(length_frame, text="Study Length:").pack(side=tk.LEFT)
        self.study_length = tk.StringVar()
        self.length_combo = ttk.Combobox(length_frame, textvariable=self.study_length,
                                        values=["1", "2", "3", "4", "5", "6", "7"],
                                        width=5, state="readonly")
        self.length_combo.pack(side=tk.LEFT, padx=5)
        self.length_combo.bind('<<ComboboxSelected>>', self.update_study_days)
        ToolTip(self.length_combo, "Number of days to process for each site")

        # Study Days
        days_frame = ttk.LabelFrame(study_frame, text="Study Days (optional)", 
                                   style='Header.TLabelframe')
        days_frame.pack(side=tk.LEFT, padx=20)
        self.day_vars = {}
        self.day_checkbuttons = {}
        for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']:
            self.day_vars[day] = tk.BooleanVar()
            self.day_checkbuttons[day] = ttk.Checkbutton(
                days_frame, text=day, variable=self.day_vars[day],
                command=self.update_action_buttons
            )
            self.day_checkbuttons[day].pack(side=tk.LEFT, padx=5)

        # Action Buttons
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(pady=2)  # Reduced vertical padding

        self.scan_button = ttk.Button(button_frame, text="Scan Files",
                                     command=self.scan_project, width=20,
                                     state='disabled')
        self.scan_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.scan_button, "Scan project folder for EC0 files -- Must load in a Project Tracker URL and Project Folder and set Study Length and Days")

        self.process_button = ttk.Button(button_frame, text="Process Files",
                                        command=self.start_processing, width=20,
                                        state='disabled')
        self.process_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.process_button, "Begin processing selected files")

        self.pause_button = ttk.Button(button_frame, text="Pause Processing",
                                      command=self.pause_processing, width=20,
                                      state='disabled')
        self.pause_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.pause_button, "Pause/Resume processing")

        self.stop_button = ttk.Button(button_frame, text="Stop Processing",
                                     command=self.stop_processing, width=20,
                                     state='disabled')
        self.stop_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.stop_button, "Stop all processing")
        
        self.help_button = ttk.Button(button_frame, text="Help Guide",
                                     command=self.show_help, width=20)
        self.help_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.help_button, "View help documentation")

        self.creature_button = ttk.Button(button_frame, text="🐈",
                                        command=self.show_creature, width=3)
        self.creature_button.pack(side=tk.LEFT, padx=5)
        ToolTip(self.creature_button, "🐈")                      

        # Bottom paned window for status and site parameters
        bottom_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        bottom_paned.grid(row=1, column=0, sticky='nsew', padx=5, pady=2)  # Reduced pady

        # Status section with reduced width
        status_frame = ttk.LabelFrame(bottom_paned, text="Status", 
                                    style='Header.TLabelframe')
        self.status = tk.Text(status_frame, width=66)
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status.yview)
        self.status.configure(yscrollcommand=scrollbar.set)
        self.status.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        bottom_paned.add(status_frame, weight=1)

        # Site Parameters section
        site_frame = ttk.LabelFrame(bottom_paned, text="Site Parameters", 
                                   style='Header.TLabelframe')
        self.config_frame = SiteConfigFrame(site_frame, study_length=self.study_length, process_type=self.process_type)
        self.config_frame.pack(fill=tk.BOTH, expand=True)
        bottom_paned.add(site_frame, weight=5)

        # Footer with credits
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=(0, 2))

        credits_text = "Created by Mark Salhany • Questions or comments or ideas? Contact: mark.salhany@idaxdata.com"
        credits_label = ttk.Label(footer_frame, 
                                 text=credits_text,
                                 font=('TkDefaultFont', 8),
                                 foreground='gray')
        credits_label.pack(side=tk.RIGHT, padx=5)

        # Configure text widget tags for formatting
        self.status.tag_configure('warning', foreground='red', font=('TkDefaultFont', 9, 'bold'))
        self.status.tag_configure('site_title', font=('TkDefaultFont', 9, 'bold'))

        # Status tag configuration is moved here since self.status now exists
        self.status.tag_configure('warning', foreground='red', font=('TkDefaultFont', 9, 'bold'))
        self.status.tag_configure('site_title', font=('TkDefaultFont', 9, 'bold'))    

    
    def update_study_days(self, *args):
        """Update day selection based on study length"""
        try:
            study_length = self.study_length.get()

            # Reset all days first if not in Bellevue mode
            if self.process_type.get() != ProcessingType.BELLEVUE:
                for day, var in self.day_vars.items():
                    var.set(False)
                    button = self.day_checkbuttons[day]
                    button.state(['!disabled'])

            if study_length == "7" or self.process_type.get() == ProcessingType.BELLEVUE:
                # Check and disable all days
                for day, var in self.day_vars.items():
                    var.set(True)
                    button = self.day_checkbuttons[day]
                    button.state(['disabled'])
            elif study_length == "3":
                # Auto-select Tue, Wed, Thu for 3-day studies
                self.day_vars['Tue'].set(True)
                self.day_vars['Wed'].set(True)
                self.day_vars['Thu'].set(True)

            # Update button states
            self.update_action_buttons()

        except Exception as e:
            self.logger.error(f"Error updating study days: {e}")

    def on_process_type_change(self, *args):
        """Handle changes to processing type selection"""
        proc_type = self.process_type.get()
        self.update_status(f"Active processing type: {proc_type}", tag='processing')

        if proc_type == ProcessingType.BELLEVUE:
            # Force study length to 7 and disable changes
            self.study_length.set("7")
            self.length_combo.state(['disabled'])
            
            # Check and disable all day checkboxes
            for day, var in self.day_vars.items():
                var.set(True)
                self.day_checkbuttons[day].state(['disabled'])
            
            # Force study length to 7 in site parameters if they exist
            if hasattr(self, 'config_frame') and hasattr(self.config_frame, 'site_vars'):
                for site_num in self.config_frame.site_vars:
                    self.config_frame.site_vars[site_num]['study_length'].set("7")
                    if site_num in self.config_frame.entries:
                        self.config_frame.entries[site_num]['study_length'].state(['disabled'])
        else:
            # Re-enable study length selection
            self.length_combo.state(['!disabled'])
            
            # Re-enable day checkboxes
            for day, var in self.day_vars.items():
                self.day_checkbuttons[day].state(['!disabled'])
                
            # Re-enable study length in site parameters
            if hasattr(self, 'config_frame') and hasattr(self.config_frame, 'site_vars'):
                for site_num in self.config_frame.entries:
                    self.config_frame.entries[site_num]['study_length'].state(['!disabled'])
                    
        # Make sure to update the buttons state
        self.update_action_buttons()

    def update_action_buttons(self):
        try:
            study_length_selected = bool(self.study_length.get())
            sheet_loaded = self.project_data is not None
            folder_selected = bool(self.folder_path.get())
            
            enabled = all([
                study_length_selected,
                sheet_loaded,
                folder_selected
            ])

            if hasattr(self, 'scan_button'):
                self.scan_button.configure(state='normal' if enabled else 'disabled')

            if hasattr(self, 'process_button'):
                self.process_button.configure(state='normal' if enabled else 'disabled')

        except Exception as e:
            self.logger.error(f"Error updating buttons: {e}")

    def show_help(self):
       """Display help window with usage instructions"""
       help_window = tk.Toplevel(self.root)
       help_window.title("IDAX Tube Data Processor - Help Guide")
       help_window.geometry("800x600")

       # Add a proper button frame at the bottom
       button_frame = ttk.Frame(help_window)
       button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
       
       ttk.Button(button_frame, text="Close", 
                  command=help_window.destroy).pack(side=tk.RIGHT, padx=5)

       # Create a frame with scrollbar
       frame = ttk.Frame(help_window)
       frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

       # Add scrollbar
       scrollbar = ttk.Scrollbar(frame)
       scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

       # Create text widget
       text = tk.Text(frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
       text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
       scrollbar.config(command=text.yview)

       # Configure tags for formatting
       text.tag_configure('header', font=('TkDefaultFont', 12, 'bold'))
       text.tag_configure('subheader', font=('TkDefaultFont', 10, 'bold'))
       text.tag_configure('normal', font=('TkDefaultFont', 10))

       # Add help content
       help_content = r"""
    IDAX Tube Data Processor Help Guide

1. Initial Setup
   --------------
   First Launch Configuration:
   - When you run the application for the first time, you'll be prompted to set the path to MTExec.exe.
     Navigate to the location where MTExec.exe is installed (usually under C:\Program Files\MetroCount\MTExec\MTExec.exe).
   - You'll also be prompted to enter the Working Tracker URL.
     Provide the Google Sheets URL used for tracking your projects.

   Subsequent Launches:
   - These settings are saved, so you won't need to enter them again unless you want to change them via the Settings menu.

2. Project Structure
   ------------------
   Folder Requirements:
   - Project folder must contain an 'ec0' subfolder
   - Inside 'ec0', folders containing the ec0 files must contain a date and certain keywords: 'pickup', 'pu', or 'pups'
   - Example pickup folder names: '10-9 pups', '10-21 pickup', 'Oct21_pu'
   - Program will save .txt files to default location: C:\Users\YOURNAME\Documents\MetroCount\MTE 5.06\Output

3. Google Sheet Requirements
   --------------------------
   Required Headers:
   - Map # - Site identifier
   - Site # - Secondary identifier
   - STREET SET - Main street name
   - REFERENCE ST - Cross street or reference
   - Status - Site processing status
   - PROCESSING NAMING - Output file naming
   - NB/EB TUBE - North/East bound box number
   - SB/WB TUBE - South/West bound box number
   - Date Set - Date tubes were set
   - DATE P/U - Pickup date
   - One Way - Checkbox for one-way streets
   - Speed Limit - Required for Cities/Speed processing - Fill in speed limits on Google Sheet for each site if required

4. Processing Types
   -----------------
   Standard Processing:
   - Basic tube data processing without addl speed stats
   - Uses 15M_ALL STATS profile

   Cities/Speed Processing:
   - Includes extra speed information compared against posted speed limit
   - Requires speed limits to be enterred for each site
   - Uses 15M_ALL STATS SPEED STATS profile

   Bellevue Processing:
   - Produces two .txt files per site: 7-day processing plus 7-day with Friday to Monday masked out
   - Study length is automatically set to 7 days and cannot be changed
   - All study days are selected and cannot be modified

5. One-Way Streets
   ----------------
   - Check the One-Way box for sites that are one-way streets
   - Sites that only have a SB/WB box must be marked as one-way
   - Direction is determined by which field (NB/EB or SB/WB) contains the box number
   - Sites with only SB/WB box and no one-way check will be skipped during processing

6. Site Parameters
   ----------------
   - Processing Name: Name used for output files
   - Study Length: Duration of study (1-7 days, automatically set to 7 for Bellevue processing)
   - Study Days (optional): Select the days to include in the study (not required for scanning files, automatically set for Bellevue processing)
   - Begin/End Dates: Study period dates
   - Speed Limit: Required for Cities/Speed processing

7. Processing Operations
   ----------------------
   Scan Files:
   - Scans project folder for matching EC0 files
   - Validates site configurations
   - Reports matched and unmatched files
   - Study days selection is now optional for scanning files (except for Bellevue processing)

   Process Files:
   - Processes each site according to configuration
   - Can be paused/stopped during operation
   - Updates site status upon completion
   - For Bellevue processing, generates two .txt files per site with different date ranges

8. Common Issues
   --------------
   - Missing Box Numbers: Ensure box numbers match the tracker
   - One-Way Configuration: Sites with single box must be marked as one-way
   - Speed Limits: Required for Cities/Speed processing
   - File Matching: EC0 files must be in pickup folders
   - Sheet Headers: Must match exactly as listed above

9. Tips
   -------
   - For best results, close or minimize other windows before starting processing
   - Use 'Set All' for begin dates to align multiple sites
   - Check warnings for misconfigured one-way sites
   - Verify study length matches project requirements (automatically set for Bellevue processing)
   - Keep MTExec window visible during processing
   - Don't move mouse during processing

For additional support, questions, or issues:
Contact Mark Salhany at mark.salhany@idaxdata.com
    """

       # Insert content with formatting
       text.insert('end', help_content)

       # Make text readonly
       text.config(state='disabled')

       # Center window on screen
       help_window.update_idletasks()
       width = help_window.winfo_width()
       height = help_window.winfo_height()
       x = (help_window.winfo_screenwidth() // 2) - (width // 2)
       y = (help_window.winfo_screenheight() // 2) - (height // 2)
       help_window.geometry(f'+{x}+{y}')

       # Make window stay on top
       help_window.transient(self.root)
       help_window.grab_set()    

    def load_working_tracker(self):
        """Set the sheet URL to the working tracker and load the sheet data"""
        # Get URL from settings manager
        working_tracker_url = self.settings_manager.get_working_tracker_url()
        self.sheet_url.set(working_tracker_url)
        
        # Call the method to load the sheet data
        self.load_sheet_data()
    

    def load_sheet_data(self):
        """Load and parse data from Google Sheet with improved column handling and validation"""
        try:
            sheet_url = self.sheet_url.get()
            if not sheet_url:
                messagebox.showerror("Error", "Please enter a Google Sheet URL")
                return

            self.update_status("Loading sheet data...")

            # Setup Google Sheets API
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
            creds = service_account.Credentials.from_service_account_file(
                'client_secrets.json', scopes=SCOPES)

            sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
            if not sheet_id_match:
                raise ValueError("Invalid Google Sheet URL")
            sheet_id = sheet_id_match.group(1)
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()

            # Store sheet ID and name for later use
            self.sheet_id = sheet_id

            # Get the sheet metadata to extract the sheet name
            spreadsheet_metadata = sheet.get(
                spreadsheetId=sheet_id,
                fields='properties/title'
            ).execute()
            self.sheet_name.set(spreadsheet_metadata['properties']['title'])

            # Define required columns and their possible variations
            REQUIRED_COLUMNS = {
                'Map #': ['Map #', 'MAP #', 'Map#', 'MAP#', 'MAP'],
                'Site #': ['Site #', 'SITE #', 'Site#', 'SITE#', 'SITE'],
                'STREET SET': ['STREET SET', 'Street Set', 'STREET'],
                'REFERENCE ST': ['REFERENCE ST', 'Reference St', 'REF ST', 'REFERENCE STREET'],
                'Status': ['Status', 'STATUS'],
                'PROCESSING NAMING': ['PROCESSING NAMING', 'Processing Naming', 'PROC NAME'],
                'NB/EB TUBE': ['NB/EB TUBE', 'NB/EB', 'NB TUBE', 'EB TUBE'],
                'SB/WB TUBE': ['SB/WB TUBE', 'SB/WB', 'SB TUBE', 'WB TUBE'],
                'Date Set': ['Date Set', 'DATE SET', 'SET DATE'],
                'DATE P/U': ['DATE P/U', 'Date P/U', 'PICKUP DATE', 'DATE PICKUP'],
                'One Way': ['One Way', 'ONE WAY', 'ONEWAY'],
                'Speed Limit': ['Speed Limit', 'SPEED LIMIT', 'Posted Speed', 'POSTED SPEED']
            }

            # Get all data to find header row
            result = sheet.values().get(
                spreadsheetId=sheet_id,
                range='A1:Z50'  # Extended range to catch headers further to the right
            ).execute()
            values = result.get('values', [])

            # Find header row by looking for required columns
            header_row = None
            for idx, row in enumerate(values):
                # Convert row to uppercase for case-insensitive matching
                row_upper = [str(cell).upper() if cell else '' for cell in row]
                # Check if any of our required columns are present
                if any(any(variant.upper() in row_upper for variant in variants)
                       for variants in REQUIRED_COLUMNS.values()):
                    header_row = idx
                    break

            if header_row is None:
                raise ValueError("Could not find header row with required columns")

            # Get all data including headers
            result = sheet.values().get(
                spreadsheetId=sheet_id,
                range=f'A{header_row + 1}:Z1000'
            ).execute()
            values = result.get('values', [])
            if not values:
                raise ValueError("No data found in sheet")

            headers = values[0]
            data = values[1:]

            # Ensure consistent column count
            max_cols = len(headers)
            data = [row + [''] * (max_cols - len(row)) for row in data]

            # Create column mapping
            column_indices = {}
            for target_col, variants in REQUIRED_COLUMNS.items():
                found = False
                for idx, header in enumerate(headers):
                    if any(variant.upper() == str(header).upper() for variant in variants):
                        column_indices[target_col] = idx
                        found = True
                        break
                if not found and target_col != 'Speed Limit':  # Speed Limit is optional
                    raise ValueError(f"Required column '{target_col}' not found in sheet")

            # Create DataFrame with found columns
            df_dict = {}
            for target_col, idx in column_indices.items():
                df_dict[target_col] = [row[idx] if idx < len(row) else '' for row in data]

            # Create the DataFrame
            self.project_data = pd.DataFrame(df_dict)

            # Convert to proper boolean values with strict checking
            self.project_data['One Way'] = self.project_data['One Way'].apply(
                lambda x: True if pd.notna(x) and str(x).strip().upper() in ['TRUE', '1', 'YES', 'CHECKED', '✓', 'X'] else False
)

            # After creating the DataFrame but before the conversion
            #self.update_status(f"Debug - One Way values from sheet: {self.project_data['One Way'].unique()}")

            # Then try this more explicit conversion
            #self.project_data['One Way'] = self.project_data['One Way'].apply(lambda x: True if str(x).lower().strip() in ['true', '1', 'yes', 'checked', '✓'] else False)

            # Add Row Index column (actual row number in the sheet)
            self.project_data['Row Index'] = self.project_data.index + header_row + 2

            # Store sheet ID for later use
            self.sheet_id = sheet_id

            # Clean the data
            for col in self.project_data.columns:
                self.project_data[col] = self.project_data[col].fillna('').astype(str).str.strip()

            # Filter out rows without Map numbers
            self.project_data = self.project_data[self.project_data['Map #'].str.strip() != '']

            # Create full street names
            self.project_data['Full Street'] = self.project_data.apply(
                lambda row: f"{row['STREET SET']} {row['REFERENCE ST']}".strip(),
                axis=1
            )

            # Display one-way warnings
            one_way_warnings = []
            for _, row in self.project_data.iterrows():
                nb_box = str(row['NB/EB TUBE']).strip()
                sb_box = str(row['SB/WB TUBE']).strip()
                one_way = row['One Way']
                
                if not nb_box and sb_box and not one_way:
                    one_way_warnings.append(f"Map #{row['Map #']} - Possible one-way road detected but 'One Way' is not checked")

            if one_way_warnings:
                warning_text = "\n".join(one_way_warnings)
                self.status.insert('end', "\n")  # Add a line break for spacing
                self.status.insert('end', "One-Way Warnings:\n", 'warning_header')
                self.status.insert('end', warning_text + "\n", 'warning')
                self.status.see('end')

            # Display loading results
            total_sites = len(self.project_data)
            needs_processing = len(self.project_data[self.project_data['Status'] == 'NEEDS TXT FILE'])
            
            self.update_status(f"\nSheet loaded successfully.")
            self.update_status(f"Total sites found: {total_sites}")
            self.update_status(f"Sites ready to process: {needs_processing}")

            self.update_action_buttons()

        except Exception as e:
            self.logger.error(f"Error loading sheet: {str(e)}")
            self.update_status(f"Error loading sheet: {str(e)}")
            messagebox.showerror("Error", f"Failed to load sheet: {str(e)}")

    def normalize_box_number(self, box_num):
        """Normalize box number for comparison"""
        if not box_num or pd.isna(box_num):
            return ''

        # Convert to string and clean up
        box_num = str(box_num).strip().upper()

        # Remove any non-alphanumeric characters
        box_num = re.sub(r'[^A-Z0-9]', '', box_num)

        # Handle D-numbers
        if box_num.startswith('D'):
            return box_num

        # Remove leading zeros and convert back to string
        try:
            return str(int(box_num))
        except ValueError:
            return box_num

    def match_box_number(self, ec0_filename):
        """Match EC0 file to site info with improved twin set handling"""
        try:
            # Extract box number from filename
            match = re.search(r'BOX\s+(\d+|D\d+)\s+0\s+', ec0_filename, re.IGNORECASE)
            if not match:
                self.update_status(f"Could not extract box number from filename: {ec0_filename}")
                return None

            box_num = self.normalize_box_number(match.group(1))
            self.update_status(f"\nAttempting to match box {box_num}")

            needs_processing = self.project_data[self.project_data['Status'] == 'NEEDS TXT FILE']

            for _, site in needs_processing.iterrows():
                nb_box = self.normalize_box_number(str(site['NB/EB TUBE']))
                sb_box = self.normalize_box_number(str(site['SB/WB TUBE']))

                # Match box to either tube position
                if box_num in [nb_box, sb_box]:
                    self.update_status(f"Found match: Site #{site['Site #']} - {site['Full Street']}")
                    return {
                        'site': site['Site #'],
                        'street': site['Full Street'],
                        'site_info': site,
                        'is_twin_set': bool(nb_box and sb_box),
                        'primary_box': nb_box,
                        'twin_box': sb_box if box_num == nb_box else nb_box
                    }

            self.update_status(f"No match found for box {box_num}")
            return None

        except Exception as e:
            self.logger.error(f"Error matching box number: {e}")
            self.update_status(f"Error while matching: {str(e)}")
            return None

    def parse_date(self, date_str):
        """Parse date string from various formats"""
        if not date_str or pd.isna(date_str):
            return None

        try:
            # Try common date formats
            formats = [
                "%m/%d/%y",
                "%m/%d/%Y",
                "%Y-%m-%d",
                "%m-%d-%y",
                "%m-%d-%Y"
            ]

            # Extract just the date part if there's a comma (e.g., "10/21, 1730")
            if ',' in date_str:
                date_str = date_str.split(',')[0].strip()

            # Try each format
            for fmt in formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue

            return None

        except Exception as e:
            self.logger.error(f"Error parsing date {date_str}: {e}")
            return None

    def parse_folder_date(self, folder_name):
        """Parse date from pickup folder name"""
        try:
            # Handle simple format like "10-9 pups"
            match = re.search(r'(\d{1,2})-(\d{1,2})', folder_name)
            if match:
                month = int(match.group(1))
                day = int(match.group(2))
                year = datetime.now().year  # Current year
                return datetime(year, month, day)

            # Extract numbers from folder name (original logic as backup)
            nums = ''.join(c for c in folder_name if c.isdigit())
            
            if len(nums) >= 8:  # YYYYMMDD
                return datetime.strptime(nums[:8], "%Y%m%d")
            elif len(nums) >= 6:  # MMDDYY
                return datetime.strptime(nums[:6], "%m%d%y")
            elif len(nums) == 4:  # MMDD
                return datetime(datetime.now().year, int(nums[:2]), int(nums[2:]))
                
            return None

        except Exception as e:
            self.logger.error(f"Error parsing folder date '{folder_name}': {e}")
            return None

    def parse_ec0_filename(self, filename):
        """Parse EC0 filename to extract box number and pickup date"""
        try:
            # Parse format: "BOX 344 0 2024-10-18 1555"
            match = re.match(r'BOX\s+(\d+|D\d+)\s+0\s+(\d{4}-\d{2}-\d{2})\s+\d+', filename)
            if match:
                box_num = self.normalize_box_number(match.group(1))
                pickup_date = datetime.strptime(match.group(2), '%Y-%m-%d')
                return box_num, pickup_date
            return None, None
        except Exception as e:
            self.logger.error(f"Error parsing EC0 filename {filename}: {e}")
            return None, None

    def find_similar_numbers(self, target_num, all_numbers, threshold=10):
        """Find numbers within threshold distance of target"""
        try:
            target = int(target_num)
            similar = []
            for num in all_numbers:
                try:
                    current = int(num)
                    if current != target and abs(current - target) <= threshold:
                        similar.append(current)
                except ValueError:
                    continue
            return sorted(similar)
        except ValueError:
            return []

    def extract_box_number(self, filename):
        """Extract box number from EC0 filename"""
        # Try both formats
        patterns = [
            r'BOX\s+(\d+|D\d+)\s+0\s+',  # With BOX prefix
            r'^(\d+|D\d+)\s+0\s+'         # Without BOX prefix
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename, re.IGNORECASE)
            if match:
                return self.normalize_box_number(match.group(1))
        
        return None

    def match_files_to_site(self, site_info, pickup_folders, ec0_path):
        """Match EC0 files to a site based on box numbers and dates - handles nested folder structure"""
        try:
            map_num = site_info.get('Map #', '').strip()
            nb_box = self.normalize_box_number(str(site_info.get('NB/EB TUBE', '')).strip())
            sb_box = self.normalize_box_number(str(site_info.get('SB/WB TUBE', '')).strip())
            set_date = site_info.get('Date Set', '')
            stated_pickup = site_info.get('DATE P/U', '')

            # Calculate expected pickup date
            if stated_pickup and stated_pickup.strip():
                try:
                    expected_pickup = datetime.strptime(stated_pickup, "%m/%d/%Y")
                except:
                    expected_pickup = datetime.strptime(set_date, "%m/%d/%Y") + timedelta(days=int(self.study_length.get()) + 1)
            else:
                expected_pickup = datetime.strptime(set_date, "%m/%d/%Y") + timedelta(days=int(self.study_length.get()) + 1)

            # Calculate date range for matching (±1 day from expected pickup)
            date_range = [
                expected_pickup - timedelta(days=1),
                expected_pickup + timedelta(days=1)
            ]

            best_match = None
            for pickup_folder in pickup_folders:
                folder_path = os.path.join(ec0_path, pickup_folder) if isinstance(pickup_folder, str) else pickup_folder
                
                # Check folder date
                folder_date = self.parse_folder_date(os.path.basename(folder_path))
                if not folder_date:
                    continue
                    
                # Check if folder date is within range
                if not (date_range[0] <= folder_date <= date_range[1]):
                    continue

                ec0_files = [f for f in os.listdir(folder_path) if f.upper().endswith('.EC0')]
                
                # Track both tubes for twin sets
                nb_file = None
                sb_file = None
                
                # Find matching files for both tubes if present
                for ec0_file in ec0_files:
                    if not ec0_file.upper().endswith('.EC0'):
                        continue
                        
                    # Try both formats - with or without BOX prefix
                    box_match = re.search(r'(?:BOX\s+)?(\d+|D\d+)\s+0\s+', ec0_file, re.IGNORECASE)
                    if box_match:
                        box_num = self.normalize_box_number(box_match.group(1).strip())
                        
                        if box_num == nb_box:
                            nb_file = {
                                'file': ec0_file,
                                'path': os.path.join(folder_path, ec0_file),
                                'pickup_folder': folder_path
                            }
                        elif box_num == sb_box:
                            sb_file = {
                                'file': ec0_file,
                                'path': os.path.join(folder_path, ec0_file),
                                'pickup_folder': folder_path
                            }

                # Create match if we found at least one tube
                if nb_file or sb_file:
                    # Determine if it's a twin set
                    is_twin = bool(nb_box and sb_box and nb_file and sb_file)
                    
                    match = {
                        'site': map_num,
                        'street': f"{site_info.get('STREET SET', '')} {site_info.get('REFERENCE ST', '')}".strip(),
                        'site_info': site_info,
                        'primary_file': nb_file['file'] if nb_file else sb_file['file'],
                        'primary_path': nb_file['path'] if nb_file else sb_file['path'],
                        'pickup_folder': os.path.basename(nb_file['pickup_folder'] if nb_file else sb_file['pickup_folder']),
                        'is_twin_set': is_twin,
                        'twin_file': sb_file['file'] if is_twin else None,
                        'twin_path': sb_file['path'] if is_twin else None,
                        'study_length': self.study_length.get()
                    }
                    best_match = match
                    break

                               

            return best_match

        except Exception as e:
            self.logger.error(f"Error matching files for site {site_info.get('Map #')}: {e}")
            self.update_status(f"Error in matching: {str(e)}")
            return None

    def infer_pickup_date(self, site_info):
        """Infer pickup date based on set date and study length"""
        try:
            set_date_str = site_info[SheetColumns.SET_DATE]
            if pd.isna(set_date_str):
                return None

            # Parse set date
            set_date = None
            if ',' in set_date_str:  # Format: "10/21, 1730, jb"
                date_part = set_date_str.split(',')[0].strip()
                set_date = datetime.strptime(f"{date_part}/2024", "%m/%d/%Y")
            else:
                try:
                    set_date = datetime.strptime(set_date_str, "%m/%d/%Y")
                except:
                    try:
                        set_date = datetime.strptime(set_date_str, "%m/%d/%y")
                    except:
                        return None

            # Calculate pickup date based on study length
            study_length = int(self.study_length.get())
            if study_length == 7:
                # For 7-day studies, pickup is 8 days after set date
                pickup_date = set_date + timedelta(days=8)
            else:
                # For other studies, calculate based on selected days
                study_start = set_date + timedelta(days=1)  # Start day after set
                pickup_date = study_start + timedelta(days=study_length)

            return pickup_date

        except Exception as e:
            self.logger.error(f"Error inferring pickup date: {e}")
            return None

    def browse_folder(self):
        """Open folder browser dialog"""
        folder = filedialog.askdirectory(title="Select Project Folder")
        if folder:
            self.folder_path.set(folder)
            self.project_name.set(os.path.basename(folder.rstrip('/\\')))
            self.update_status(f"Selected project folder: {folder}")
            self.update_action_buttons()


    def update_status(self, message, tag=None, end='\n'):
        """Update status display with optional formatting"""
        self.status.insert('end', message)
        if tag:
            # Apply tag to just the newly inserted text
            last_line_start = self.status.index("end-1c linestart")
            last_line_end = self.status.index("end-1c")
            self.status.tag_add(tag, last_line_start, last_line_end)
        if end:
            self.status.insert('end', end)
        self.status.see('end')
        self.root.update()
        self.logger.info(message)

    def get_site_info(self, ec0_filename):
        """Get site information from project data based on EC0 filename"""
        if self.project_data is None:
            return None

        try:
            # Extract box number from EC0 filename (e.g., "BOX D51" -> "D51")
            match = re.search(r'BOX\s+([A-Za-z0-9]+)', ec0_filename, re.IGNORECASE)
            if match:
                box_num = self.normalize_box_number(match.group(1))

                # Case-insensitive matching for tube columns
                matches = self.project_data[
                    (self.project_data[SheetColumns.NB_TUBE].apply(self.normalize_box_number) == box_num) |
                    (self.project_data[SheetColumns.SB_TUBE].apply(self.normalize_box_number) == box_num)
                ]

                if not matches.empty:
                    return matches.iloc[0]

                self.update_status(f"No match found for box {box_num}")

        except Exception as e:
            self.logger.error(f"Error parsing EC0 filename {ec0_filename}: {e}")

        return None

    def parse_pickup_date(self, date_str):
        """Parse pickup date from various possible formats"""
        try:
            if not date_str or pd.isna(date_str):
                return None
                
            date_str = str(date_str).strip()
            
            # Try various date formats
            formats = [
                "%m/%d/%Y",    # 10/15/2024
                "%m/%d/%y",    # 10/15/24
                "%Y-%m-%d",    # 2024-10-15
                "%m/%d",       # 10/15
            ]
            
            for fmt in formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    if fmt == "%m/%d":  # Add current year if not provided
                        current_year = datetime.now().year
                        parsed_date = parsed_date.replace(year=current_year)
                    return parsed_date
                except ValueError:
                    continue
                    
            return None
        except Exception as e:
            self.logger.error(f"Error parsing pickup date '{date_str}': {e}")
            return None
            
    def calculate_pickup_date(self, set_date_str, study_length):
        """Calculate expected pickup date based on set date and study length"""
        try:
            if not set_date_str:
                return None
                
            set_date = self.parse_pickup_date(set_date_str)
            if not set_date:
                return None
                
            # Pickup occurs study_length + 1 days after set date
            return set_date + timedelta(days=int(study_length) + 1)
        except Exception as e:
            self.logger.error(f"Error calculating pickup date: {e}")
            return None        

    def sort_map_numbers(self, map_num):
        """Convert map number to sortable integer"""
        try:
            # Extract only the digits from the map number
            num = ''.join(filter(str.isdigit, str(map_num)))
            return int(num) if num else float('inf')
        except:
            return float('inf')

    def find_similar_numbers(self, target_num, folder_path, threshold=10):
        """Find similar box numbers within specific pickup folder"""
        try:
            target = int(target_num)
            similar = []
            
            # Only look at EC0 files in the specified folder
            ec0_files = [f for f in os.listdir(folder_path) if f.upper().endswith('.EC0')]
            
            # Get box numbers from this folder only
            folder_boxes = set()
            for file in ec0_files:
                box_num = self.extract_box_number(file)
                if box_num:
                    try:
                        box_int = int(box_num)
                        folder_boxes.add(box_int)
                    except ValueError:
                        continue
            
            # Find similar numbers only from this folder
            for num in folder_boxes:
                if num != target and abs(num - target) <= threshold:
                    similar.append(num)
                    
            return sorted(similar)
        except ValueError:
            return []    

    def scan_project(self):
        """Scan project folder for EC0 files with flexible subfolder support"""
        if not self.study_length.get():
            messagebox.showwarning("Warning", "Please set study length before scanning files")
            return

        if self.project_data is None:
            messagebox.showerror("Error", "Please load sheet data first")
            return

        if not self.folder_path.get():
            messagebox.showerror("Error", "Please select a project folder")
            return

        missing_box_warnings = []
        duplicate_box_warnings = []

        self.update_status("\nScanning project structure...")

        ec0_path = os.path.join(self.folder_path.get(), "ec0")
        if not os.path.exists(ec0_path):
            self.update_status("No ec0 folder found!")
            return

        # Find all pickup folders recursively
        def find_pickup_folders(root_path):
            pickup_folders = []
            for dirpath, dirnames, _ in os.walk(root_path):
                # Check if this directory is a pickup folder
                if any(x in os.path.basename(dirpath).lower() for x in ["pickup", "pu", "pups"]):
                    # Get the path between ec0 and pickup folder as the group name
                    rel_path = os.path.relpath(dirpath, root_path)
                    parent_path = os.path.dirname(rel_path)
                    group_name = parent_path if parent_path != "." else "Main"
                    
                    pickup_folders.append({
                        'path': dirpath,
                        'group': group_name,
                        'name': os.path.basename(dirpath)
                    })
            return pickup_folders

        # Get all pickup folders
        pickup_folders = find_pickup_folders(ec0_path)

        if not pickup_folders:
            self.update_status("No pickup folders found!")
            return

        # Group pickup folders for status reporting
        groups = {}
        for folder in pickup_folders:
            if folder['group'] not in groups:
                groups[folder['group']] = []
            groups[folder['group']].append(folder)

        # Report found folders by group
        self.update_status("\nPickup folders found:")
        if len(groups) > 1:  # Only show group headers if there are multiple groups
            for group_name, folders in groups.items():
                group_display = group_name if group_name != "Main" else "Main Directory"
                self.update_status(f"\n{group_display}:", tag='site_title')
                for folder in folders:
                    self.update_status(f"  • {folder['name']}")
        else:
            # If only one group, just list the folders directly
            for folder in next(iter(groups.values())):
                self.update_status(f"  • {folder['name']}")

        # Find matching sites
        needs_processing = self.project_data[self.project_data['Status'] == 'NEEDS TXT FILE']
        needs_processing = needs_processing.sort_values(by='Map #', key=lambda x: x.map(self.sort_map_numbers))
        
        # Track matched boxes to avoid suggesting them as similar matches
        matched_boxes = set()
        
        # Sites matching status
        self.update_status("\nAttempting to match ec0 files to sites:", tag='header')
        self.update_status("")
        for _, site in needs_processing.iterrows():
            map_num = site.get('Map #', '').strip()
            nb_box = self.normalize_box_number(str(site.get('NB/EB TUBE', '')).strip())
            sb_box = self.normalize_box_number(str(site.get('SB/WB TUBE', '')).strip())
            
            # no box num check
            if not sb_box and not nb_box:
                missing_box_warnings.append(f"Map #{map_num} has no box numbers")
            
            # duplicate box num check:
            if nb_box and sb_box and nb_box == sb_box:
                duplicate_box_warnings.append(
                    f"Map #{map_num} has the same box number ({nb_box}) in both NB/EB and SB/WB fields"
                )
                      
            is_twin = bool(nb_box and sb_box)
            box_text = f"(Boxes {nb_box}/{sb_box})" if is_twin else f"(Box {nb_box or sb_box})"
            set_type = "Twin set" if is_twin else "Single set"
            
            expected_pickup = None
            set_date = self.parse_date(site.get('Date Set', ''))
            if set_date:
                expected_pickup = set_date + timedelta(days=int(self.study_length.get()) + 1)
                pickup_folder = next((folder['name'] for folder in pickup_folders 
                    if abs((self.parse_folder_date(folder['name']) - expected_pickup).days) <= 1), None)
                folder_text = f". Searching in {pickup_folder}..."
            else:
                folder_text = ". No valid set date found."
                
            self.status.insert('end', f"Map #{site['Map #']}", ('bold',))
            self.update_status(f" - {set_type} {box_text}{folder_text}")

        # Initialize lists for tracking matches
        fully_matched_sites = []
        partially_matched_sites = []
        unmatched_sites = []

        # Search through all sites
        for _, site in needs_processing.iterrows():
            map_num = site.get('Map #', '').strip()
            nb_box = self.normalize_box_number(str(site.get('NB/EB TUBE', '')).strip())
            sb_box = self.normalize_box_number(str(site.get('SB/WB TUBE', '')).strip())
            site_title = f"Map #{map_num} - {site.get('STREET SET', '')} {site.get('REFERENCE ST', '')}".strip()

            # Search all pickup folders for matches
            best_match = None
            for folder in pickup_folders:
                current_match = self.match_files_to_site(site, [folder['path']], ec0_path)
                if current_match:
                    current_match['group'] = folder['group']
                    current_match['pickup_folder_name'] = folder['name']
                    if not best_match or (
                        current_match['is_twin_set'] and not best_match['is_twin_set']
                    ):
                        best_match = current_match

            if best_match:
                if best_match['is_twin_set'] or not (nb_box and sb_box):
                    # Full match or single intended tube site matched - add to fully_matched_sites
                    fully_matched_sites.append({
                        'site': map_num,
                        'street': site_title,
                        'group': best_match['group'],
                        'site_info': site,
                        **best_match
                    })
                else:
                    # It's a twin site but we only found one tube - add to partially_matched_sites only
                    partially_matched_sites.append({
                        'site': map_num,
                        'title': site_title,
                        'street': site_title,
                        'group': best_match['group'],
                        'site_info': site,
                        **best_match
                    })
            else:
                unmatched_sites.append({
                    'title': site_title,
                    'nb_box': nb_box,
                    'sb_box': sb_box
                })

        # Display results grouped by folder group
        self.update_status("\nMATCHED SITES:", tag='success_header')
        self.update_status("")
        if len(groups) > 1:  # Only show group headers if there are multiple groups
            for group_name in sorted(groups.keys()):
                group_matches = [site for site in fully_matched_sites if site['group'] == group_name]
                if group_matches:
                    group_display = group_name if group_name != "Main" else "Main Directory"
                    self.update_status(f"\n{group_display}:", tag='site_title')
                    for site in group_matches:
                        # Determine if single or double set
                        set_type = "(twin set)" if site['is_twin_set'] else "(single set)"
                        
                        # Display site with bold Map # and set type
                        site_base = f"  Map #{site['site']}"
                        site_name = f" - {site['street']} {set_type}"
                        
                        self.status.insert('end', site_base, ('site_title', 'bold'))
                        self.status.insert('end', site_name)
                        self.status.insert('end', '\n')
                        
                        if site['is_twin_set']:
                            self.update_status(f"    NB/EB: {os.path.basename(site['primary_file'])} — in {site['pickup_folder_name']}")
                            self.update_status(f"    SB/WB: {os.path.basename(site['twin_file'])} — in {site['pickup_folder_name']}")
                        else:
                            self.update_status(f"    Found: {os.path.basename(site['primary_file'])} — in {site['pickup_folder_name']}")
                        
                        # Add blank line between sites
                        self.update_status("")
        else:
            # If only one group, just show the sites directly
            for site in fully_matched_sites:
                # Display site with bold Map # and set type
                set_type = "(twin set)" if site['is_twin_set'] else "(single set)"
                site_base = f"  Map #{site['site']}"
                site_name = f" - {site['street']} {set_type}"
                
                self.status.insert('end', site_base, ('site_title', 'bold'))
                self.status.insert('end', site_name)
                self.status.insert('end', '\n')
                
                if site['is_twin_set']:
                    self.update_status(f"    NB/EB: {os.path.basename(site['primary_file'])} — in {site['pickup_folder_name']}")
                    self.update_status(f"    SB/WB: {os.path.basename(site['twin_file'])} — in {site['pickup_folder_name']}")
                else:
                    self.update_status(f"    Found: {os.path.basename(site['primary_file'])} — in {site['pickup_folder_name']}")
                
                # Add blank line between sites
                self.update_status("")
        
        # Partially matched sites
        if partially_matched_sites:
            self.update_status("\nPARTIALLY MATCHED SITES:", tag='partial_header')
            for site in partially_matched_sites:
                self.update_status("")  # Add space between entries
                self.status.insert('end', f"  Map #{site['site']}", ('bold',))
                self.update_status(f" - {site['street']}")
                
                # Get the relevant pickup folder where we found the matching file
                pickup_folder = os.path.dirname(site['primary_path'])
                
                # Determine which box was found and which is missing
                nb_box = self.normalize_box_number(str(site['site_info'].get('NB/EB TUBE', '')))
                sb_box = self.normalize_box_number(str(site['site_info'].get('SB/WB TUBE', '')))
                
                if site['primary_file']:
                    found_box = self.extract_box_number(site['primary_file'])
                    if found_box == nb_box:
                        self.update_status(f"    Found EC0 file for NB/EB - Box {nb_box}")
                        self.update_status(f"    Missing EC0 file for SB/WB - Box {sb_box}")
                        # Only look for unmatched similar boxes in the same folder
                        if sb_box not in matched_boxes:  # Only if the missing box isn't matched elsewhere
                            similar_matches = self.find_similar_box_numbers(sb_box, pickup_folder, matched_boxes)
                            if similar_matches:
                                self.update_status(f"    Similar box numbers found in {site['pickup_folder_name']}:", tag='info')
                                for match in similar_matches[:3]:
                                    self.update_status(f"      • Box {match['box']} ({match['file']})", tag='info')
                    else:
                        self.update_status(f"    Found EC0 file for SB/WB - Box {sb_box}")
                        self.update_status(f"    Missing EC0 file for NB/EB - Box {nb_box}")
                        # Only look for unmatched similar boxes in the same folder
                        if nb_box not in matched_boxes:  # Only if the missing box isn't matched elsewhere
                            similar_matches = self.find_similar_box_numbers(nb_box, pickup_folder, matched_boxes)
                            if similar_matches:
                                self.update_status(f"    Similar box numbers found in {site['pickup_folder_name']}:", tag='info')
                                for match in similar_matches[:3]:
                                    self.update_status(f"      • Box {match['box']} ({match['file']})", tag='info')

        # Show unmatched sites      
        if unmatched_sites:
            self.update_status("\nUNMATCHED SITES:", tag='warning_header')
            for site in unmatched_sites:
                self.update_status("")  # Add space between entries
                self.status.insert('end', f"  {site['title']}", ('bold',))
                self.update_status("")
                if site['nb_box'] or site['sb_box']:
                    if site['nb_box']:
                        self.update_status(f"    Missing EC0 file for NB/EB - Box {site['nb_box']}")
                    if site['sb_box']:
                        self.update_status(f"    Missing EC0 file for SB/WB - Box {site['sb_box']}")
                    # Check for similar box numbers in all pickup folders
                    for folder in pickup_folders:
                        if site['nb_box']:
                            similar_matches = self.find_similar_box_numbers(site['nb_box'], folder['path'], matched_boxes)
                            if similar_matches:
                                self.update_status(f"    Similar box numbers found in {folder['name']}:", tag='info')
                                for match in similar_matches[:3]:
                                    self.update_status(f"      • Box {match['box']} ({match['file']})", tag='info')

                        if site['sb_box']:
                            similar_matches = self.find_similar_box_numbers(site['sb_box'], folder['path'], matched_boxes)
                            if similar_matches:
                                self.update_status(f"    Similar box numbers found in {folder['name']}:", tag='info')
                                for match in similar_matches[:3]:
                                    self.update_status(f"      • Box {match['box']} ({match['file']})", tag='info')

        # Show additional warnings
        if missing_box_warnings or duplicate_box_warnings:
            self.update_status("\nADDITIONAL WARNINGS:", tag='warning_header')
            
            if missing_box_warnings:
                self.update_status("")
                self.update_status("Sites with no box numbers:", tag='warning')
                for warning in missing_box_warnings:
                    self.update_status(f"  • {warning}", tag='warning')
                    
            if duplicate_box_warnings:
                self.update_status("")
                self.update_status("Sites with duplicate box numbers:", tag='warning')
                for warning in duplicate_box_warnings:
                    self.update_status(f"  • {warning}", tag='warning')  
                    
        # Store matched files and update site parameters
        self.ec0_files = fully_matched_sites
        if fully_matched_sites:
            self.update_status("\nLoading Site Parameters...")
            self.update_status(f"Number of sites loaded: {len(fully_matched_sites)}")
            self.config_frame.populate_files(self.ec0_files)
        else:
            self.update_status("No files matched for processing.")

    def click_element(self, image_name, confidence=0.9, timeout=10, message="", grayscale=False, clicks=2):
        """Find and click an element using image recognition"""
        try:
            # Only show messages for specific important actions
            important_actions = {
                'save_button.png': "Saving file...",
                'add_files.png': "Loading files...",
                'custom_list.png': "Preparing report...",
                'speed_stats.png': "Using speed stats profile",
                'all_stats.png': "Using standard profile"
            }
            
            if image_name in important_actions:
                self.update_status(important_actions[image_name], tag='processing')

            start_time = time.time()

            while time.time() - start_time < timeout:
                # Force MTExec window to front
                mtexec_windows = pyautogui.getWindowsWithTitle("MTExec")
                if mtexec_windows:
                    mtexec_windows[0].activate()
                    time.sleep(0.2)

                # Try to find the element
                location = pyautogui.locateOnScreen(f'images/{image_name}', confidence=confidence, grayscale=grayscale)
                if location:
                    x, y = pyautogui.center(location)
                    # Move and click
                    pyautogui.moveTo(x, y, duration=0.1)
                    time.sleep(0.1)
                    pyautogui.click(clicks=clicks)
                    time.sleep(0.1)
                    return True

                time.sleep(0.1)

            if image_name in important_actions:
                self.update_status(f"Failed to {important_actions[image_name].lower()}", tag='warning')
            return False

        except Exception as e:
            self.logger.error(f"Error clicking {image_name}: {e}")
            return False


    def double_click_element(self, image_name, confidence=0.9, timeout=10, message="", grayscale=False):
        """Find and double-click an element"""
        self.update_status(f"Looking for {message if message else image_name}...")
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(f'images/{image_name}', confidence=confidence, grayscale=grayscale)
                if location:
                    x, y = pyautogui.center(location)
                    pyautogui.doubleClick(x, y)
                    self.update_status(f"Double-clicked {message if message else image_name}")
                    time.sleep(0.1)
                    return True
            except Exception as e:
                self.logger.error(f"Error finding {image_name}: {e}")
                time.sleep(0.1)

        self.update_status(f"Could not find {image_name}")
        return False

    def adjust_date_field(self, field_type, target_date, reference_date):
        """
        Adjust the date in the Time Range dialog by clicking left/right arrows under the middle date box.
        field_type: 'begin' or 'end'
        target_date: The date to set
        reference_date: The date from which to calculate the difference
        """
        try:
            # Calculate the number of days difference
            days_diff = (target_date - reference_date).days

            # Determine the direction
            direction = 'left' if days_diff < 0 else 'right'
            arrow_image = f'{direction}_arrow.png'  # 'left_arrow.png' or 'right_arrow.png'

            # Click the arrow the required number of times
            for _ in range(abs(days_diff)):
                if not self.click_element(arrow_image, message=f"{direction.capitalize()} arrow", confidence=0.8):
                    raise Exception(f"Could not click {direction} arrow")
                time.sleep(0.2)
        except Exception as e:
            self.logger.error(f"Error adjusting {field_type} date: {e}")
            raise

    def store_site_parameters(self):
        """Store all parameters from the site parameters frame before processing"""
        try:
            self.stored_parameters = {}
            for site_num, vars_dict in self.config_frame.site_vars.items():
                self.stored_parameters[site_num] = {
                    'study_length': vars_dict['study_length'].get(),
                    'begin_date': vars_dict['begin_date'].get(),
                    'end_date': vars_dict['end_date'].get(),
                    'speed_limit': vars_dict['speed_limit'].get(),
                    'proc_name': vars_dict['proc_name'].get(),
                    'one_way': vars_dict['one_way'].get(),
                    'nb_box': vars_dict['nb_box'].get(),
                    'sb_box': vars_dict['sb_box'].get(),
                    'set_date_parsed': vars_dict['set_date_parsed']
                }
            self.update_status("Stored updated site parameters")
        except Exception as e:
            self.logger.error(f"Error storing site parameters: {e}")
            raise    

    def read_date_from_box(self):
        """Read date from currently highlighted box using clipboard"""
        try:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)
            date_text = pyperclip.paste().strip()
            self.update_status(f"Read date: {date_text}")
            return date_text
        except Exception as e:
            self.logger.error(f"Error reading date from box: {e}")
            return None
   
    def adjust_time_range(self, current_params):
        """Adjust time range by reading and modifying dates"""
        try:
            self.update_status("\nStarting time range adjustment...")
            
            # Store initial position (Auto-wrap dropdown)
            initial_x, initial_y = pyautogui.position()
            
            # ---- BEGIN DATE SECTION ----
            self.update_status("\nAdjusting begin date...")
            
            # Move to begin date text box
            pyautogui.moveTo(
                initial_x - 275,
                initial_y - 100,
                duration=0.2
            )
            pyautogui.click()
            time.sleep(0.2)
            
            # Read and adjust begin date
            begin_date = self.read_date_from_box()
            if not begin_date:
                raise Exception("Could not read begin date")
                
            target_begin = current_params['begin_date']
            if not target_begin:
                raise Exception("No target begin date found in parameters")
                
            days_diff = self.calculate_date_adjustment(begin_date, target_begin)
            
            # Track where we end up after begin date adjustment
            mouse_position = "text_box"  # Default position
            
            if days_diff != 0:
                # Move to appropriate arrow
                if days_diff < 0:
                    pyautogui.moveRel(-20, 25)  # Left arrow
                    mouse_position = "left_arrow"
                else:
                    pyautogui.moveRel(20, 25)   # Right arrow
                    mouse_position = "right_arrow"
                
                # Click required number of times
                for _ in range(abs(days_diff)):
                    pyautogui.click()
                    time.sleep(0.1)
            
            # ---- END DATE SECTION ----
            self.update_status("\nMoving to end date section...")
            
            # Move to end date text box based on current position
            if mouse_position == "text_box":
                pyautogui.moveRel(250, 0)  # From begin date text box
            elif mouse_position == "left_arrow":
                pyautogui.moveRel(290, -25)  # From left arrow
            elif mouse_position == "right_arrow":
                pyautogui.moveRel(250, -25)  # From right arrow
            
            pyautogui.click()
            time.sleep(0.2)
            
            # Read and adjust end date
            end_date = self.read_date_from_box()
            if not end_date:
                raise Exception("Could not read end date")
                
            target_end = current_params['end_date']
            if not target_end:
                raise Exception("No target end date found in parameters")
                
            days_diff = self.calculate_date_adjustment(end_date, target_end)
            
            if days_diff != 0:
                # Move to arrows under end date - THIS IS THE FIXED PART
                if days_diff < 0:
                    # Move to left arrow (green circled)
                    pyautogui.moveRel(-15, 28)
                else:
                    # Move to right arrow (blue circled)
                    pyautogui.moveRel(15, 28)
                
                # Click required number of times
                for _ in range(abs(days_diff)):
                    pyautogui.click()
                    time.sleep(0.1)
            
            # Accept changes
            pyautogui.press('enter')
            self.update_status("Time range adjustment complete!")
            return True
            
        except Exception as e:
            self.logger.error(f"Error adjusting time range: {e}")
            self.update_status(f"Time range adjustment error: {str(e)}")
            return False

    def attempt_save(self, proc_name, max_attempts=2):
        """Attempt to save file with retry logic"""
        try:
            for attempt in range(max_attempts):
                # Click save button
                if not self.click_element('save_button.png', message="Save button", confidence=0.8, clicks=1):
                    raise Exception("Could not find Save button")
                time.sleep(0.2)
                
                # Clear existing text and write new name
                current_name = proc_name if attempt == 0 else f"{proc_name}_{attempt+1}"
                pyautogui.hotkey('ctrl', 'a')  # Select all
                time.sleep(0.1)
                pyautogui.write(current_name)
                time.sleep(0.2)
                pyautogui.press('enter')
                time.sleep(1.0)  # Wait for potential overwrite dialog
                
                # Check for overwrite dialog
                try:
                    if pyautogui.locateOnScreen('images/confirm_save.png', confidence=0.8):
                        self.update_status("Handling file overwrite...")
                        # Focus the window
                        confirm_windows = pyautogui.getWindowsWithTitle("Confirm Save As")
                        if confirm_windows:
                            confirm_windows[0].activate()
                        time.sleep(0.2)
                        
                        # Try clicking Yes button
                        if self.click_element('yes.png', message="Yes button", confidence=0.8, clicks=1):
                            time.sleep(0.5)
                            
                            # Check if save was successful
                            if not pyautogui.locateOnScreen('images/save_report.png', confidence=0.8):
                                self.update_status("Save successful!", tag='success')
                                return True
                        
                        # If we're still here, save failed
                        if attempt < max_attempts - 1:
                            self.update_status("Save failed, retrying with new name...")
                            continue
                    else:
                        # No overwrite needed
                        self.update_status("Save successful!")
                        return True
                        
                except Exception as e:
                    self.logger.error(f"Error during save attempt {attempt+1}: {e}")
                    if attempt < max_attempts - 1:
                        continue
                    
            return False
            
        except Exception as e:
            self.logger.error(f"Error in save attempt: {e}")
            return False

    def process_ec0_files(self):
        """Process EC0 files through MTExec"""       
        # Optimize environment before processing
        try:
            # Disable screen saver
            pyautogui.press('f15')  # Prevents screen saver
            # Set process priority
            import psutil
            p = psutil.Process()
            p.nice(psutil.HIGH_PRIORITY_CLASS)
        except:
            pass
        
        if not os.path.exists(self.mtexec_path):
            messagebox.showerror("Error", 
                "MTExec.exe not found at configured path. Please check Settings > Set MTExec Path")
            return

        failed_sites = []
        skipped_sites = []
        successful_sites = []
        
        
        if not self.ec0_files:
            messagebox.showerror("Error", "No files to process. Please scan first.")
            return

        try:            
            # Get current parameters from site parameters frame
            site_params = self.config_frame.get_site_data()
            self.update_status("\nStarting EC0 file processing...")

            # Launch MTExec at start
            subprocess.Popen(self.mtexec_path)
            time.sleep(0.5)

            last_error = False  # Track if previous site had error
            for file_info in self.ec0_files:
                # Check for stop event
                if self.stop_event.is_set():
                    self.update_status("Processing stopped by user.")
                    break

                # Check for pause event
                while self.pause_event.is_set():
                    time.sleep(0.1)
                    if self.stop_event.is_set():
                        self.update_status("Processing stopped by user during pause.")
                        break

                if self.stop_event.is_set():
                    break

                site_num = file_info['site']
                current_params = site_params[site_num]

                # Skip if site is disabled
                if site_num in self.config_frame.disabled_sites:
                    reason = "Missing speed limit" if self.process_type.get() == ProcessingType.CITIES_SPEED else "Invalid one-way configuration"
                    if self.process_type.get() == ProcessingType.CITIES_SPEED:
                        # Use warning tag for speed limit messages
                        self.update_status(f"\nSkipping Site #{site_num} - {reason}", tag='warning')
                    else:
                        self.update_status(f"\nSkipping Site #{site_num} - {reason}")
                    skipped_sites.append((site_num, reason))
                    continue
                
                self.update_status(f"\nProcessing Site #{site_num} - {file_info['street']}", tag='site_title')
                self.update_status(f"Study length: {current_params['study_length']} days", tag='info')
                if file_info['is_twin_set']:
                    self.update_status("Processing as twin set")

                try:
                    # If last site had error, ensure MTExec is running
                    if last_error:
                        self.update_status("Restarting MTExec after previous error...")
                        for window in pyautogui.getWindowsWithTitle("MTExec"):
                            window.close()
                        time.sleep(0.7)
                        subprocess.Popen(self.mtexec_path)
                        time.sleep(0.7)

                    # Force focus on MTExec window
                    mtexec_windows = pyautogui.getWindowsWithTitle("MTExec")
                    if mtexec_windows:
                        mtexec_windows[0].activate()
                        time.sleep(0.2)

                    # Initial setup
                    if not all([
                        self.click_element('new_report.png', message="New Report button"),
                        self.click_element('add_files.png', message="Add Files button")
                    ]):
                        raise Exception("Failed to initialize MTExec")

                    time.sleep(0.2)

                    # Navigate to correct folder
                    pickup_folder_path = os.path.dirname(file_info['primary_path'])
                    pyautogui.hotkey('alt', 'd')
                    time.sleep(0.1)
                    pyautogui.write(pickup_folder_path)
                    time.sleep(0.1)
                    pyautogui.press('enter')
                    time.sleep(0.2)

                    # Load EC0 files
                    if not self.click_element('load_data_filename.png', message="File name entry", clicks=1):
                        raise Exception("Could not find file name entry")
                                            
                    # Prepare EC0 files based on twin set status
                    if file_info['is_twin_set'] and file_info['twin_file']:
                        file_names = f'"{file_info["primary_file"]}" "{file_info["twin_file"]}"'
                        self.update_status("Loading twin set files:", tag='info')
                        self.update_status(f"  • {file_info['primary_file']}", tag='processing')
                        self.update_status(f"  • {file_info['twin_file']}", tag='processing')
                    else:
                        file_names = f'"{file_info["primary_file"]}"'
                        self.update_status(f"Loading single file: {file_info['primary_file']}", tag='info')

                    pyautogui.write(file_names)
                    time.sleep(0.1)
                    
                    # Look for the Open button instead of pressing tab
                    if self.click_element('open_ec0.png', message="Open button", confidence=0.8):
                        time.sleep(0.1)
                    else:
                        # Fallback to tab method if button not found                        
                        pyautogui.press('enter')
                        time.sleep(0.2)
                                           
                    pyautogui.press('tab')
                    time.sleep(0.1)
                    pyautogui.press('enter')
                    time.sleep(0.15)

                    # Double-click "Custom List Report"
                    if not self.double_click_element('custom_list.png', message="Custom List Report"):
                        raise Exception("Could not find Custom List Report")
                    time.sleep(0.2)

                    cities_speed = self.process_type.get() == ProcessingType.CITIES_SPEED
                    is_one_way = current_params['one_way']

                    # Handle one-way configuration first if needed
                    if is_one_way:                        
                        
                        # Determine direction based on box presence
                        has_nb = bool(current_params['nb_box'].strip())
                        has_sb = bool(current_params['sb_box'].strip())
                        direction = "NB/EB" if has_nb else "SB/WB"
                        
                        self.update_status(f"Processing as one-way {direction}", tag='info')
                        
                        # Navigate to Direction tab
                        pyautogui.press('right', presses=3, interval=0.1)
                        pyautogui.press('enter')
                        time.sleep(0.1)
                        
                        # Click direction modifier dropdown
                        if not self.click_element('direction_mod.png', message="Direction modifier dropdown", clicks=1):
                            raise Exception("Could not find direction modifier dropdown")
                        time.sleep(0.1)
                        
                        # Set direction
                        if not has_nb:  # SB/WB one-way
                            pyautogui.press('pagedown')
                            time.sleep(0.1)
                            pyautogui.press('enter')
                        else:  # NB/EB one-way
                            pyautogui.press('pagedown')
                            time.sleep(0.1)
                            pyautogui.press('up')
                            time.sleep(0.1)
                            pyautogui.press('enter')
                        
                        # Confirm changes
                        pyautogui.press('enter')
                        time.sleep(0.1)

                    # Handle speed limit setup if needed
                    if cities_speed:
                        self.update_status(f"Speed limit set to: {current_params['speed_limit']}", tag='info')
                        
                        # Click Advanced button
                        if not self.click_element('advanced.png', message="Advanced button", clicks=1):
                            raise Exception("Could not find Advanced button")
                        time.sleep(0.1)
                        
                        # Click Speed tab
                        if not self.click_element('speed_tab.png', message="Speed tab", clicks=1):
                            raise Exception("Could not find Speed tab")
                        time.sleep(0.1)
                        
                        # Find "Default speed limit" text first
                        if not self.click_element('default_speed.png', message="Default speed limit text", clicks=1):
                            raise Exception("Could not find Default speed limit text")
                        time.sleep(0.1)

                        # Move left to the text box
                        current_x, current_y = pyautogui.position()
                        pyautogui.moveTo(current_x - 78, current_y)
                        pyautogui.click()
                        time.sleep(0.1)

                        # Enter speed limit
                        speed_limit = current_params['speed_limit'].strip()
                        if not speed_limit:
                            raise Exception("No speed limit specified for processing")
                        
                        # Select all existing text and replace
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(0.1)
                        pyautogui.write(speed_limit)
                        time.sleep(0.1)
                        
                        # Click OK to close Advanced window
                        if not self.click_element('ok.png', message="OK button", clicks=1):
                            raise Exception("Could not find OK button")
                        time.sleep(0.1)
                        
                        # Navigate to time range with tab presses after speed setup
                        for _ in range(10):
                            pyautogui.press('tab')
                            time.sleep(0.05)  # Small delay between tabs
                        pyautogui.press('enter')
                        time.sleep(0.1)
                    elif is_one_way:
                        # If only one-way (no speed), navigate with one right press
                        pyautogui.press('right')
                        pyautogui.press('enter')
                    else:
                        # Standard navigation to time range
                        pyautogui.press('down', presses=4, interval=0.1)
                        pyautogui.press('enter')

                    time.sleep(0.1)

                    # Click Auto-Wrap dropdown
                    if not self.click_element('auto_wrap.png', message="Auto Wrap dropdown", clicks=1):
                        raise Exception("Could not find Auto Wrap dropdown")
                    time.sleep(0.1)

                    # Handle time range based on site's study length
                    study_length = current_params['study_length']
                    if study_length == "7":
                        # Handle 7-day studies
                        x, y = pyautogui.position()
                        pyautogui.moveTo(x, y + 325)
                        time.sleep(0.1)
                        pyautogui.click()
                        time.sleep(0.1)
                        pyautogui.press('enter')
                        time.sleep(0.1)
                    else:
                        # For non-7-day studies
                        pyautogui.press('home')  # Select "None"
                        time.sleep(0.2)
                        pyautogui.press('enter')
                        time.sleep(0.1)

                        # Adjust dates using reference points
                        if not self.adjust_time_range(current_params):
                            raise Exception("Failed to adjust time range")
                        
                        time.sleep(0.2)

                    # Navigate to report vortex
                    pyautogui.press('down', presses=2, interval=0.1)
                    pyautogui.press('enter')
                    time.sleep(0.2)

                    # Select appropriate processing profile
                    profile_image = 'speed_stats.png' if cities_speed else 'all_stats.png'
                    self.update_status(f"Using {'speed stats' if cities_speed else 'standard'} profile")

                    if not self.double_click_element(profile_image, message="Processing profile", confidence=0.8):
                        raise Exception("Could not find processing profile")

                    pyautogui.press('enter', presses=3, interval=0.1)
                    time.sleep(0.1)

                    # Save with processing name from parameters
                    if not self.click_element('save_button.png', message="Save button", confidence=0.8, clicks=1):
                        raise Exception("Could not find Save button")
                    time.sleep(0.2)

                    # Get processing name and write it
                    proc_name = current_params['proc_name'].strip()
                    if not proc_name:
                        proc_name = f"Site_{site_num}_Data"
                    self.update_status(f"Saving as: {proc_name}", tag='info')
                    pyautogui.write(proc_name)
                    time.sleep(0.2)

                    # Initial enter press after entering filename
                    pyautogui.press('enter')
                    time.sleep(0.3)

                    try:
                        # Check if we got an overwrite dialog
                        if pyautogui.getWindowsWithTitle("Confirm Save As"):
                            self.update_status("File exists, attempting to overwrite...", tag='info')
                            # Press Alt+Y to confirm overwrite
                            pyautogui.hotkey('alt', 'y')
                            time.sleep(0.2)
                            
                            # Check if we're back to Save report to file... (overwrite failed)
                            if pyautogui.getWindowsWithTitle("Save report to file"):
                                modified_name = f"{proc_name}_2"
                                self.update_status(f"Trying alternate name: {modified_name}", tag='info')
                                pyautogui.hotkey('ctrl', 'a')
                                time.sleep(0.1)
                                pyautogui.write(modified_name)
                                time.sleep(0.2)
                                pyautogui.press('enter')  # First enter after new name
                                time.sleep(0.2)
                                pyautogui.press('enter')  # Second enter to confirm save
                                time.sleep(0.2)
                        else:
                            # No overwrite needed, just hit enter again to confirm save
                            pyautogui.press('enter')
                            time.sleep(0.5)
                        
                        # Final verification
                        if pyautogui.getWindowsWithTitle("Save report to file"):
                            raise Exception("Failed to complete save operation")
                        else:
                            self.update_status("Save completed successfully", tag='success')
                            
                    except Exception as e:
                        self.logger.error(f"Save error: {str(e)}")
                        raise

                    time.sleep(0.2)  # Wait for save to complete

                    if self.process_type.get() == ProcessingType.BELLEVUE:
                        try:                           
                            time.sleep(0.2)
                            # Click Local profile button
                            if not self.click_element('kim.png', message="Local profile button", confidence=0.8):
                                raise Exception("Could not find Local profile button")
                            time.sleep(0.2)
                            
                            # Navigate to time range
                            pyautogui.press('right', presses=4, interval=0.1)
                            pyautogui.press('enter')
                            time.sleep(0.2)
                            
                            # Navigate to Use exclusion checkbox
                            pyautogui.press('right', presses=2, interval=0.1)
                            pyautogui.press('space')  # Toggle checkbox
                            time.sleep(0.1)
                            
                            # Click Mask button
                            pyautogui.press('right')
                            pyautogui.press('enter')
                            time.sleep(0.2)
                            
                            # First click Clear all
                            if not self.click_element('clear_all.png', message="Clear all button", confidence=0.8):
                                raise Exception("Could not find Clear all button")
                            time.sleep(0.2)
                            
                            # Get window reference for relative positioning
                            window = pyautogui.getActiveWindow()
                            if not window or "Exclusion" not in window.title:
                                raise Exception("Lost focus of Exclusion window")
                            
                            # Function to find and click pattern with region restriction
                            def click_day_plus(image_name, description):
                                # Define region for top section only (adjust these values if needed)
                                region = (
                                    window.left,
                                    window.top,
                                    window.left + 100,  # Only look in leftmost area
                                    window.top + 150    # Only look in top portion
                                )
                                
                                location = pyautogui.locateOnScreen(
                                    f'images/{image_name}',
                                    confidence=0.8,
                                    region=region
                                )
                                
                                if location:
                                    x, y = pyautogui.center(location)
                                    pyautogui.click(x, y)
                                    self.update_status(f"Re-enabled {description}", tag='success')
                                    time.sleep(0.2)
                                    return True
                                else:
                                    raise Exception(f"Could not find {description}")
                            
                            # Click the plus signs next to Tue, Wed, Thu in order
                            click_day_plus('tue.png', "Tuesday")
                            click_day_plus('wed.png', "Wednesday")
                            click_day_plus('thu.png', "Thursday")
                            
                            # Accept the selections
                            pyautogui.press('tab', presses=2, interval=0.1)
                            pyautogui.press('enter')
                            time.sleep(0.3)
                            pyautogui.press('tab', presses=1, interval=0.1)
                            pyautogui.press('enter')
                            time.sleep(0.1)
                            pyautogui.press('tab', presses=2, interval=0.1)
                            pyautogui.press('enter')
                            time.sleep(0.2)
                                                                              
                            # Process through to save with modified name
                            time.sleep(0.2)
                            if not self.click_element('save_button.png', message="Save button", confidence=0.8):
                                raise Exception("Could not find Save button")
                            
                            # Determine filename with .1 suffix
                            base_name = proc_name
                            if "_2" in proc_name:
                                modified_name = f"{base_name}.1"
                            else:
                                modified_name = f"{base_name}.1"
                                
                            # Attempt save
                            pyautogui.hotkey('ctrl', 'a')  # Select all existing text
                            pyautogui.write(modified_name)
                            time.sleep(0.2)
                            pyautogui.press('enter')
                            # Initial enter press after entering filename
                            pyautogui.press('enter')
                            time.sleep(0.3)
                                                                                  
                            # Handle potential overwrite
                            time.sleep(0.3)
                            if pyautogui.locateOnScreen('images/yes.png', confidence=0.8):
                                modified_name = f"{base_name}_2.1"
                                pyautogui.hotkey('ctrl', 'a')
                                pyautogui.write(modified_name)
                                time.sleep(0.2)
                                pyautogui.press('enter')
                            
                            pyautogui.press('enter')  # Second enter to confirm save
                            time.sleep(0.2)
                            
                            self.update_status(f"Completed Bellevue processing for {base_name}", tag='success')
                            
                        except Exception as e:
                            self.logger.error(f"Error in Bellevue additional processing: {e}")                            

                    # Update site status
                    if 'site_info' in file_info and 'Row Index' in file_info['site_info']:
                        row_index = int(file_info['site_info']['Row Index'])
                        self.update_status(f"Updated status for Map #{file_info['site']}", tag='success')
                        self.update_site_status(row_index, 'READY TO PROCESS')
                    else:
                        self.update_status(f"Warning: Missing row information for Map #{file_info['site']}")

                    successful_sites.append({
                        'site': site_num,
                        'street': file_info['street'],
                        'files': [file_info['primary_file']]
                    })
                    if file_info['is_twin_set'] and file_info['twin_file']:
                        successful_sites[-1]['files'].append(file_info['twin_file'])

                    last_error = False

                except Exception as e:
                    last_error = True
                    error_msg = str(e)                    
                    self.update_status(f"Error processing site {site_num}: {error_msg}", tag='warning')
                    failed_sites.append((site_num, error_msg))
                    
                    if not self.handle_error_with_timeout(site_num, error_msg):
                        break

            self.update_status("\nProcessing complete!")

        except Exception as e:
            self.logger.error(f"Processing error: {str(e)}")
            self.update_status(f"Error during processing: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

        finally:
            # Report results
            self.update_status("\n=== PROCESSING RESULTS ===\n", tag='bold')
            
            # Show successful sites in green
            if successful_sites:
                self.update_status("Successfully Processed Sites:", tag='success')
                for site in successful_sites:
                    self.update_status(f"  Map #{site['site']} - {site['street']}", tag='success')
                    for file in site['files']:
                        self.update_status(f"    • {file}", tag='success')
                self.update_status("")
            
            # Show failed sites in red
            if failed_sites:
                self.update_status("Failed Sites:", tag='warning')
                for site_num, error in failed_sites:
                    self.update_status(f"  Map #{site_num}: {error}", tag='warning')
                self.update_status("")
            
            # Show skipped sites
            if skipped_sites:
                self.update_status("Skipped Sites:")
                for site_num, reason in skipped_sites:
                    self.update_status(f"  Map #{site_num}: {reason}")
            
            # Reset buttons
            self.process_button.state(['!disabled'])
            self.pause_button.state(['disabled'])
            self.stop_button.state(['!disabled'])
            self.pause_button.config(text="Pause Processing")         
       
    def handle_error_with_timeout(self, site_num, error_msg):
        """Handle processing error with 10-second timeout for user response"""
        response = [False]
        event = threading.Event()
        dialog = None
        
        def show_dialog():
            nonlocal dialog
            dialog = tk.Toplevel()
            dialog.title("Error")
            dialog.attributes('-topmost', True)
            dialog.lift()
            dialog.focus_force()  # Force focus
            
             # Make dialog wider and adjust minimum height
            dialog.geometry("500x200")  # Increased width and height
            dialog.minsize(500, 200)    # Set minimum size
            
            # Center the dialog
            x = dialog.winfo_screenwidth()//2 - 250  # Adjusted for new width
            y = dialog.winfo_screenheight()//2 - 100  # Adjusted for new height
            dialog.geometry(f"+{x}+{y}")
            
            # Move mouse to neutral position
            neutral_x = dialog.winfo_screenwidth()//2
            neutral_y = dialog.winfo_screenheight()//2 + 200
            pyautogui.moveTo(neutral_x, neutral_y, duration=0.2)
            
            msg_frame = ttk.Frame(dialog, padding="10")
            msg_frame.pack(fill='x', expand=True)
            
            msg_label = ttk.Label(msg_frame, 
                                text=f"Error processing site {site_num}:\n{error_msg}\nContinue with remaining sites?",
                                wraplength=350)
            msg_label.pack(pady=5)
            
            timer_label = ttk.Label(msg_frame, 
                                  text="10", 
                                  foreground='red',
                                  font=('TkDefaultFont', 12, 'bold'))
            timer_label.pack(pady=5)
            
            btn_frame = ttk.Frame(dialog)
            btn_frame.pack(fill='x', pady=10)
            
            def on_yes():
                response[0] = True
                event.set()
                dialog.destroy()
                
            def on_no():
                response[0] = False
                event.set()
                dialog.destroy()
                
            yes_btn = ttk.Button(btn_frame, text="Yes", command=on_yes)
            no_btn = ttk.Button(btn_frame, text="No", command=on_no)
            yes_btn.pack(side='left', expand=True, padx=5)
            no_btn.pack(side='right', expand=True, padx=5)
            
            for i in range(9, -1, -1):
                if event.is_set():
                    break
                dialog.after(1000 * (10-i-1), lambda count=i: timer_label.config(text=str(count)))
            
            def auto_close():
                if not event.is_set():
                    response[0] = True
                    event.set()
                    dialog.destroy()
                    
            dialog.after(10000, auto_close)
        
        dialog_thread = threading.Thread(target=show_dialog)
        dialog_thread.daemon = True
        dialog_thread.start()
        
        event.wait(timeout=10)
        
        # Close any open MTExec windows
        pyautogui.press('escape')  # Close any open dialogs
        time.sleep(0.2)
        for window in pyautogui.getWindowsWithTitle("MTExec"):
            window.close()
        time.sleep(0.7)
        
        if response[0]:  # Only restart MTExec if continuing
            self.update_status("Continuing with next site...")
            subprocess.Popen(self.mtexec_path)
            time.sleep(0.7)
        else:
            self.update_status("Processing stopped by user.")
        
        return response[0]

    def calculate_date_adjustment(self, current_date_str, target_date_str):
        """Calculate days difference between current and target dates"""
        try:
            # Parse current date (format: "Tue 22")
            current_parts = current_date_str.split()
            current_day = int(current_parts[1])
            
            # Parse target date (format: "Tue 22")
            target_parts = target_date_str.split()
            target_day = int(target_parts[1])
            
            # Calculate difference
            days_diff = target_day - current_day
            
            if days_diff == 0:
                self.update_status(f"Date adjustment: Current date ({current_date_str}) matches target!")
            else:
                direction = "Adding" if days_diff > 0 else "Subtracting"
                self.update_status(f"Date adjustment: {current_date_str} -> {target_date_str} ({direction} {abs(days_diff)} days)")
            return days_diff
            
        except Exception as e:
            self.logger.error(f"Error calculating date adjustment: {e}")
            return 0      

    def start_processing(self):
        """Start processing in a separate thread"""
        # Reset events
        self.pause_event.clear()
        self.stop_event.clear()

        # Disable buttons appropriately
        self.process_button.state(['disabled'])
        self.pause_button.state(['!disabled'])
        self.stop_button.state(['!disabled'])

        # Start processing in a new thread
        self.processing_thread = threading.Thread(target=self.process_ec0_files)
        self.processing_thread.start()

    def pause_processing(self):
        """Toggle pause state of processing"""
        if not self.pause_event.is_set():
            self.pause_event.set()
            self.update_status("Processing paused.")
            self.pause_button.config(text="Resume Processing")
        else:
            self.pause_event.clear()
            self.update_status("Processing resumed.")
            self.pause_button.config(text="Pause Processing")

    def stop_processing(self):
        """Stop the ongoing processing"""
        self.stop_event.set()
        self.update_status("Stopping processing...")
        # Optionally, wait for the thread to finish
        if self.processing_thread:
            self.processing_thread.join()
        # Reset buttons
        self.process_button.state(['!disabled'])
        self.pause_button.state(['disabled'])
        self.stop_button.state(['disabled'])
        self.pause_button.config(text="Pause Processing")


        # Check project folder
        project_folder = self.folder_path.get()
        if project_folder:
            self.update_status(f"\nProject Folder Configuration:")
            self.update_status(f"- Base path: {project_folder}")

            # Check EC0 folder
            ec0_path = os.path.join(project_folder, "ec0")
            if os.path.exists(ec0_path):
                self.update_status("- EC0 folder: Found")

                # Check pickup folders
                pickup_folders = [f for f in os.listdir(ec0_path)
                                if any(x in f.lower() for x in ["pickup", "pu", "pups"])]
                if pickup_folders:
                    self.update_status(f"- Pickup folders found: {len(pickup_folders)}")
                else:
                    self.update_status("Warning: No pickup folders found")
            else:
                self.update_status("Warning: EC0 folder not found")
        else:
            self.update_status("Warning: No project folder selected")

        # Check processing configuration
        self.update_status(f"\nProcessing Configuration:")
        self.update_status(f"- Processing type: {self.process_type.get()}")
        self.update_status(f"- Study length: {self.study_length.get()} days")

        selected_days = [day for day, var in self.day_vars.items() if var.get()]
        if selected_days:
            self.update_status(f"- Selected days: {', '.join(selected_days)}")
        else:
            self.update_status("Warning: No study days selected")

        # Check MTExec installation
        if os.path.exists(self.mtexec_path):
            self.update_status("\nMTExec: Installed and found")
        else:
            self.update_status("\nWarning: MTExec not found at expected path")

        # Check images folder
        required_images = [
            'new_report.png',
            'add_files.png',
            'time_range.png',
            'auto_wrap.png',
            'last_seven_aligned_days.png',
            'none_option.png',
            'begin_date_field.png',
            'end_date_field.png',
            'ok_button.png',
            'next_button.png',
            'advanced_button.png',
            'speed_tab.png',
            'default_speed_field.png',
            '15m_all_stats.png',
            '15m_all_stats_speed_stats.png',
            'yes_button.png',
            'save_button.png',
            'file_name_field.png',
            'open_button.png',
            'custom_list.png'
        ]

        missing_images = []
        for image in required_images:
            if not os.path.exists(os.path.join('images', image)):
                missing_images.append(image)

        if missing_images:
            self.update_status("\nWarning: Missing required images:")
            for image in missing_images:
                self.update_status(f"- {image}")
        else:
            self.update_status("\nAll required images found")

        self.update_status("\nTest mode complete - No files were processed")

    def update_site_status(self, row_index, new_status):
        """Update the status of a site in the Google Sheet"""
        try:
            # Setup Google Sheets API first
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            creds = service_account.Credentials.from_service_account_file(
                'client_secrets.json', scopes=SCOPES)
            service = build('sheets', 'v4', credentials=creds)

            # Get a larger range to find headers
            result = service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id,
                range='A1:Z100'  # Get more rows to find headers
            ).execute()
            values = result.get('values', [])

            # Find status column by checking each row until we find it
            status_col = None
            for row in values:
                try:
                    status_col = row.index('Status')
                    break
                except (ValueError, AttributeError):
                    continue

            if status_col is None:
                raise Exception("Could not find Status column")

            # Convert to column letter
            status_col_letter = chr(65 + status_col)

            # Update the cell with correct column letter
            cell_range = f"{status_col_letter}{row_index}"
            body = {'values': [[new_status]]}
            
            result = service.spreadsheets().values().update(
                spreadsheetId=self.sheet_id,
                range=cell_range,
                valueInputOption='RAW',
                body=body
            ).execute()

        except Exception as e:
            self.logger.error(f"Error updating site status: {e}")
            self.update_status(f"Error updating site status: {e}", tag='warning')

    def get_column_letter(self, col_idx):
        """Convert zero-based column index to Excel-style column letter"""
        result = ""
        while col_idx >= 0:
            rem = col_idx % 26
            result = chr(65 + rem) + result
            col_idx = (col_idx // 26) - 1
        return result
        

    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"Application error: {str(e)}")
            messagebox.showerror("Error", f"Application error: {str(e)}")
        finally:
            # Cleanup on exit
            try:
                # Close any remaining MTExec windows
                for window in pyautogui.getWindowsWithTitle("MTExec"):
                    window.close()
                # Stop music playback
                self.is_playing = False
                mixer.music.stop()
                mixer.quit()
            except:
                pass

if __name__ == "__main__":
    try:
        app = TrafficProcessor()
        app.run()
    except Exception as e:
        logging.error(f"Failed to start application: {str(e)}")
        messagebox.showerror("Error", f"Failed to start application: {str(e)}")
