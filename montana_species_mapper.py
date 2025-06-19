import matplotlib
matplotlib.use('TkAgg')
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import geopandas as gpd
from shapely.geometry import Polygon, Point, box
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np
import os
from typing import Dict, List, Tuple, Optional
import re
import sys

# Montana County Map Generator
# This application generates county-based maps for Montana using lat/long data
# with year-based filtering to create two comparison maps

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    import sys, os
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_icon_path():
    """Get the path to the application icon"""
    try:
        # First try the .ico file
        if getattr(sys, 'frozen', False):
            # If running as compiled executable
            base_path = sys._MEIPASS
        else:
            # If running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        ico_path = os.path.join(base_path, 'app_icon.ico')
        png_path = os.path.join(base_path, 'app_icon.png')
        
        # Prefer ICO on Windows, PNG on other platforms
        if os.name == 'nt' and os.path.exists(ico_path):
            return ico_path
        elif os.path.exists(png_path):
            return png_path
        else:
            print("Warning: Icon files not found")
            return None
    except Exception as e:
        print(f"Warning: Error getting icon path: {str(e)}")
        return None

class SplashScreen:
    def __init__(self, parent):
        self.parent = parent
        self.splash = tk.Toplevel(parent)
        self.splash.title("Montana County Map Generator")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            try:
                self.splash.iconbitmap(icon_path)
            except Exception as e:
                print(f"Warning: Could not set icon for splash screen: {str(e)}")
        
        # Get screen dimensions
        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        
        # Calculate position
        width = 400
        height = 200
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.splash.geometry(f"{width}x{height}+{x}+{y}")
        self.splash.overrideredirect(True)
        self.splash.configure(bg='white')
        
        # Add loading text
        self.status_label = tk.Label(
            self.splash,
            text="Initializing...",
            bg='white',
            font=('Arial', 12)
        )
        self.status_label.pack(pady=20)
        
        # Add progress bar
        self.progress = ttk.Progressbar(
            self.splash,
            length=300,
            mode='determinate'
        )
        self.progress.pack(pady=20)
        
        self.splash.update()

    def update_status(self, message: str, progress: int = None):
        self.status_label.config(text=message)
        if progress is not None:
            self.progress['value'] = progress
        self.splash.update()

    def destroy(self):
        self.splash.destroy()

class ToastNotification:
    def __init__(self, parent):
        self.parent = parent
        
    def show_toast(self, message: str, duration: int = 3000, error: bool = False):
        toast = tk.Toplevel(self.parent)
        toast.overrideredirect(True)
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            try:
                toast.iconbitmap(icon_path)
            except Exception as e:
                print(f"Warning: Could not set icon for toast: {str(e)}")
        
        # Position toast at bottom right
        toast.geometry(f"+{self.parent.winfo_screenwidth() - 310}+{self.parent.winfo_screenheight() - 100}")
        
        # Configure toast appearance
        bg_color = '#ff4444' if error else '#44aa44'
        frame = tk.Frame(toast, bg=bg_color, padx=10, pady=5)
        frame.pack(fill='both', expand=True)
        
        tk.Label(
            frame,
            text=message,
            bg=bg_color,
            fg='white',
            wraplength=250,
            font=('Arial', 10)
        ).pack()
        
        toast.after(duration, toast.destroy)

class LoadingIndicator:
    def __init__(self, parent, message="Loading..."):
        self.parent = parent
        self.loading_window = tk.Toplevel(parent)
        self.loading_window.title("Loading")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.loading_window.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.loading_window.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for loading window: {str(e)}")
        
        # Get screen dimensions
        screen_width = self.loading_window.winfo_screenwidth()
        screen_height = self.loading_window.winfo_screenheight()
        
        # Calculate position
        width = 300
        height = 100
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.loading_window.geometry(f"{width}x{height}+{x}+{y}")
        self.loading_window.overrideredirect(True)
        self.loading_window.configure(bg='white')
        
        # Add loading text
        self.status_label = tk.Label(
            self.loading_window,
            text=message,
            bg='white',
            font=('Arial', 12)
        )
        self.status_label.pack(pady=(20, 10))
        
        # Add progress bar
        self.progress = ttk.Progressbar(
            self.loading_window,
            length=250,
            mode='indeterminate'
        )
        self.progress.pack(pady=(0, 20))
        
        # Start the progress bar
        self.progress.start(10)
        
        # Make sure the window is on top
        self.loading_window.lift()
        self.loading_window.attributes('-topmost', True)
        
        # Update the window
        self.loading_window.update()

    def update_message(self, message):
        self.status_label.config(text=message)
        self.loading_window.update()
    
    def destroy(self):
        self.progress.stop()
        self.loading_window.destroy()

class SummaryDialog:
    def __init__(self, parent, file_path, data):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("File Upload Success")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.window.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.window.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for summary dialog: {str(e)}")
        
        # Get screen dimensions and calculate center position
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        width = 600
        height = 500
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Set window geometry and properties
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.configure(bg='#ffffff')
        self.window.resizable(False, False)
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create canvas with scrollbar
        self.canvas = tk.Canvas(self.window, bg='#ffffff', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.canvas.yview)
        
        # Create main container frame that will be scrolled
        main_frame = tk.Frame(self.canvas, bg='#ffffff')
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Add main frame to canvas
        canvas_frame = self.canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Configure canvas scrolling
        def configure_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        def configure_canvas_width(event):
            self.canvas.itemconfig(canvas_frame, width=event.width)
        
        main_frame.bind('<Configure>', configure_scroll_region)
        self.canvas.bind('<Configure>', configure_canvas_width)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            if self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        self.mousewheel_binding = self.canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Bind window close event
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Content container with padding
        content_frame = tk.Frame(main_frame, bg='#ffffff', padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)
        
        # Success icon and title
        title_frame = tk.Frame(content_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(0, 20))
        
        # Success message
        tk.Label(
            title_frame,
            text="✓ File Loaded Successfully",
            font=('Arial', 18, 'bold'),
            bg='#ffffff',
            fg='#2ecc71'  # Green color
        ).pack()
        
        # File info section
        file_frame = tk.Frame(content_frame, bg='#ffffff')
        file_frame.pack(fill='x', pady=(0, 20))
        
        tk.Label(
            file_frame,
            text="FILE INFORMATION",
            font=('Arial', 12, 'bold'),
            bg='#ffffff',
            fg='#2c3e50'
        ).pack(anchor='w')
        
        ttk.Separator(file_frame, orient='horizontal').pack(fill='x', pady=(5, 10))
        
        # File statistics in a grid
        stats_frame = tk.Frame(file_frame, bg='#ffffff')
        stats_frame.pack(fill='x')
        
        stats = [
            ("File Name:", os.path.basename(file_path)),
            ("Total Records:", f"{len(data):,}"),
            ("Year Range:", f"{int(data['year'].min())} - {int(data['year'].max())}"),
            ("Unique Families:", f"{len(data['family'].unique()):,}"),
            ("Unique Genera:", f"{len(data['genus'].unique()):,}"),
            ("Unique Species:", f"{len(data['species'].unique()):,}")
        ]
        
        for i, (label, value) in enumerate(stats):
            row = i // 2
            col = i % 2
            
            stat_container = tk.Frame(stats_frame, bg='#f8f9fa', padx=10, pady=5)
            stat_container.grid(row=row, column=col, padx=5, pady=5, sticky='ew')
            
            tk.Label(
                stat_container,
                text=label,
                font=('Arial', 10),
                bg='#f8f9fa',
                fg='#7f8c8d'
            ).pack(anchor='w')
            
            tk.Label(
                stat_container,
                text=str(value),
                font=('Arial', 11, 'bold'),
                bg='#f8f9fa',
                fg='#2c3e50'
            ).pack(anchor='w')
        
        stats_frame.grid_columnconfigure(0, weight=1)
        stats_frame.grid_columnconfigure(1, weight=1)
        
        # Instructions section
        instructions_frame = tk.Frame(content_frame, bg='#ffffff')
        instructions_frame.pack(fill='x', pady=(20, 0))
        
        tk.Label(
            instructions_frame,
            text="NEXT STEPS",
            font=('Arial', 12, 'bold'),
            bg='#ffffff',
            fg='#2c3e50'
        ).pack(anchor='w')
        
        ttk.Separator(instructions_frame, orient='horizontal').pack(fill='x', pady=(5, 10))
        
        instructions = [
            "1. Select Family, Genus, and Species from the dropdowns",
            "2. Enter a year to filter the data (Map A will show data ≤ this year)",
            "3. Click 'Generate County Map' to create two comparison maps:",
            "   • Map A: Shows data up to your selected year",
            "   • Map B: Shows all available data",
            "4. Use 'Download County Map' to save as high-resolution TIFF"
        ]
        
        for instruction in instructions:
            tk.Label(
                instructions_frame,
                text=instruction,
                font=('Arial', 10),
                bg='#ffffff',
                fg='#34495e',
                justify='left',
                anchor='w'
            ).pack(anchor='w', pady=2)
        
        # Close button at bottom
        button_frame = tk.Frame(content_frame, bg='#ffffff')
        button_frame.pack(fill='x', pady=(20, 0))
        
        close_button = tk.Button(
            button_frame,
            text="Got it!",
            font=('Arial', 10, 'bold'),
            bg='#3498db',
            fg='white',
            relief='flat',
            padx=20,
            pady=5,
            command=self.on_closing,
            cursor='hand2'
        )
        close_button.pack()
        
        # Bind hover effects for the button
        close_button.bind('<Enter>', lambda e: close_button.configure(bg='#2980b9'))
        close_button.bind('<Leave>', lambda e: close_button.configure(bg='#3498db'))
        
        # Center the window and make it modal
        self.window.focus_set()
        
        # Ensure the window is on top
        self.window.lift()
        self.window.attributes('-topmost', True)
        self.window.attributes('-topmost', False)
        
        self.window.wait_window()
    
    def on_closing(self):
        """Handle window closing and cleanup"""
        # Unbind the mousewheel event
        self.canvas.unbind_all("<MouseWheel>")
        self.window.destroy()

class MainApplication:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide main window initially
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.root.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for main window: {str(e)}")
        
        # Show splash screen
        self.splash = SplashScreen(self.root)
        self.splash.update_status("Loading application...", 0)
        
        # Initialize variables
        self.excel_data = None
        self.montana_counties = None
        self.current_maps = None  # Will store both Map A and Map B
        
        # Add variables for species selection
        self.selected_family = tk.StringVar()
        self.selected_genus = tk.StringVar()
        self.selected_species = tk.StringVar()
        
        # Configure main window
        self.root.title("MontanaSpeciesMapper")
        self.root.state('zoomed')  # Start maximized
        
        # Initialize notification system
        self.toast = ToastNotification(self.root)
        
        # Set up the GUI
        self.initialize_gui()
        
        # Destroy splash screen and show main window
        self.splash.destroy()
        self.root.deiconify()

    def initialize_gui(self):
        # Configure style
        style = ttk.Style()
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white')
        style.configure('TButton', padding=5)
        
        # Main container
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Left panel (inputs)
        self.left_panel = ttk.Frame(self.main_container, style='TFrame')
        self.left_panel.pack(side='left', fill='y', padx=(0, 10))
        
        # Right panel (map display)
        self.right_panel = ttk.Frame(self.main_container, style='TFrame')
        self.right_panel.pack(side='right', fill='both', expand=True)
        
        self._setup_input_fields()
        self._setup_map_display()
        
        # Bind resize event
        self.root.bind('<Configure>', self.on_window_resize)

    def _setup_input_fields(self):
        # File selection
        ttk.Label(self.left_panel, text="Excel File:").pack(anchor='w', pady=(0, 5))
        self.file_frame = ttk.Frame(self.left_panel)
        self.file_frame.pack(fill='x', pady=(0, 20))
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(self.file_frame, textvariable=self.file_path_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(self.file_frame, text="Browse", command=self.load_excel).pack(side='right', padx=(5, 0))
        
        # Species Selection Section
        species_frame = ttk.LabelFrame(self.left_panel, text="Species Selection", padding="10")
        species_frame.pack(fill='x', pady=(0, 20))
        
        # Family
        ttk.Label(species_frame, text="Family:", style='TLabel').pack(fill='x')
        self.family_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_family, state="readonly")
        self.family_dropdown.pack(fill='x', pady=(0, 10))
        
        # Genus
        ttk.Label(species_frame, text="Genus:", style='TLabel').pack(fill='x')
        self.genus_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_genus, state="readonly")
        self.genus_dropdown.pack(fill='x', pady=(0, 10))
        
        # Species
        ttk.Label(species_frame, text="Species:", style='TLabel').pack(fill='x')
        self.species_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_species, state="readonly")
        self.species_dropdown.pack(fill='x', pady=(0, 10))
        
        # Year input instead of hexagon count
        ttk.Label(self.left_panel, text="Year:").pack(anchor='w', pady=(0, 5))
        year_frame = ttk.Frame(self.left_panel)
        year_frame.pack(fill='x', pady=(0, 20))
        
        self.year_var = tk.StringVar(value="2020")
        ttk.Entry(year_frame, textvariable=self.year_var).pack(side='left', fill='x', expand=True)
        
        # Color ranges (5 colors with grayscale scheme)
        self.color_ranges = []
        default_ranges = [
            (0, 0, "white"),
            (1, 10, "#e7e8e9"),  # Light gray
            (11, 100, "#bcbec0"),  # Medium gray
            (101, 1000, "#939598"),  # Dark gray
            (1001, float('inf'), "#231f20")  # Black
        ]
        
        ttk.Label(self.left_panel, text="Color Ranges:").pack(anchor='w', pady=(0, 5))
        
        for i, (min_val, max_val, color) in enumerate(default_ranges):
            range_frame = ttk.Frame(self.left_panel)
            range_frame.pack(fill='x', pady=(0, 10))
            
            min_var = tk.StringVar(value=str(min_val))
            max_var = tk.StringVar(value="∞" if max_val == float('inf') else str(max_val))
            color_var = tk.StringVar(value=color)
            
            ttk.Entry(range_frame, textvariable=min_var, width=8).pack(side='left', padx=(0, 5))
            ttk.Label(range_frame, text="-").pack(side='left', padx=5)
            ttk.Entry(range_frame, textvariable=max_var, width=8).pack(side='left', padx=(5, 10))
            
            color_entry = ttk.Entry(range_frame, textvariable=color_var)
            color_entry.pack(side='left', fill='x', expand=True)
            
            self.color_ranges.append((min_var, max_var, color_var))
        
        # Action buttons
        ttk.Button(self.left_panel, text="Generate County Map", command=self.generate_map).pack(fill='x', pady=(20, 5))
        ttk.Button(self.left_panel, text="Download County Map", command=self.download_map).pack(fill='x', pady=(5, 0))
        
        # Bind dropdowns
        self.family_dropdown.bind("<<ComboboxSelected>>", self.update_genus_dropdown)
        self.genus_dropdown.bind("<<ComboboxSelected>>", self.update_species_dropdown)

    def _setup_map_display(self):
        self.figure = Figure(figsize=(10, 8))
        self.ax = self.figure.add_subplot(111)
        # Remove the box from initial display
        self.ax.set_frame_on(False)
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.right_panel)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return
            
        try:
            # Show loading indicator
            loading = LoadingIndicator(self.root, "Loading Excel file...")
            
            self.excel_data = pd.read_excel(file_path)
            required_columns = ['lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', 'year']
            if not all(col in self.excel_data.columns for col in required_columns):
                loading.destroy()
                raise ValueError("Excel file must contain 'lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', and 'year' columns")
                
            self.file_path_var.set(file_path)
            
            # Process the data
            loading.update_message("Processing data...")
            for col in ["family", "genus", "species"]:
                self.excel_data[col] = self.excel_data[col].astype(str).str.strip().str.lower()
            
            # Convert year to numeric, handling any non-numeric values
            self.excel_data['year'] = pd.to_numeric(self.excel_data['year'], errors='coerce')
            
            # Get valid families (non-empty/non-null values)
            loading.update_message("Updating dropdowns...")
            valid_families = sorted(self.excel_data["family"].dropna().unique())
            valid_families = [f for f in valid_families if str(f).strip() and str(f).lower() != 'nan']  # Remove empty strings and 'nan'
            
            # Capitalize family names
            family_values = ["All"] + [f.title() for f in valid_families]
            
            # Update Family dropdown
            self.family_dropdown["values"] = family_values
            self.family_dropdown.set("Select Family")
            
            # Reset other dropdowns
            self.genus_dropdown.set("Select Genus")
            self.genus_dropdown["values"] = []
            self.species_dropdown.set("Select Species")
            self.species_dropdown["values"] = []
            
            # Load Montana counties
            loading.update_message("Loading Montana counties...")
            all_counties = gpd.read_file(resource_path("shapefiles/cb_2021_us_county_5m.shp"))
            self.montana_counties = all_counties[all_counties['STATEFP'] == '30']
            self.montana_counties = self.montana_counties.to_crs("EPSG:32100")
            
            # Bind dropdowns
            self.family_dropdown.bind("<<ComboboxSelected>>", self.update_genus_dropdown)
            self.genus_dropdown.bind("<<ComboboxSelected>>", self.update_species_dropdown)
            
            # Destroy loading indicator
            loading.destroy()
            
            # Show summary dialog
            SummaryDialog(self.root, file_path, self.excel_data)
            
            self.toast.show_toast("Excel file loaded successfully")
            
        except Exception as e:
            if 'loading' in locals():
                loading.destroy()
            self.toast.show_toast(f"Error loading file: {str(e)}", error=True)

    def dms_to_decimal(self, coord):
        """
        Convert a coordinate in DMS format (e.g., '44°41.576'') to decimal degrees.
        Handles both unicode and ascii degree/minute/second symbols.
        """
        if isinstance(coord, float) or isinstance(coord, int):
            return float(coord)
        if not isinstance(coord, str):
            return float('nan')
        # Remove unwanted characters and normalize
        coord = coord.replace("'", "'").replace("″", '"').replace("""", '"').replace(""", '"')
        dms_pattern = r"(\d+)[°\s]+(\d+(?:\.\d+)?)[\'′]?\s*(\d*(?:\.\d+)?)[\"″]?"
        match = re.match(dms_pattern, coord.strip())
        if match:
            deg = float(match.group(1))
            min_ = float(match.group(2))
            sec = float(match.group(3)) if match.group(3) else 0.0
            return deg + min_ / 60 + sec / 3600
        try:
            return float(coord)
        except Exception:
            return float('nan')

    def convert_coordinates(self, row):
        """Convert coordinates taking into account direction (N/S, E/W) and DMS/decimal formats"""
        try:
            lat = self.dms_to_decimal(row['lat'])
            long = self.dms_to_decimal(row['long'])
            
            # Convert direction values to string and handle potential NaN/float values
            lat_dir = str(row['lat_dir']).strip().upper() if pd.notna(row['lat_dir']) else 'N'
            long_dir = str(row['long_dir']).strip().upper() if pd.notna(row['long_dir']) else 'W'
            
            # Validate direction values
            if lat_dir not in ['N', 'S']:
                print(f"Invalid latitude direction: {lat_dir}, defaulting to 'N'")
                lat_dir = 'N'
            if long_dir not in ['E', 'W']:
                print(f"Invalid longitude direction: {long_dir}, defaulting to 'W'")
                long_dir = 'W'
            
            # Adjust for direction
            if lat_dir == 'S':  # If Southern hemisphere
                lat = -lat
            if long_dir == 'W':  # If Western hemisphere
                long = -long
            
            # Montana is roughly between 44°N to 49°N and 104°W to 116°W
            # Validate the coordinates are somewhat reasonable
            if not (44 <= abs(lat) <= 49 and 104 <= abs(long) <= 116):
                print(f"Warning: Coordinates ({lat}, {long}) might be outside Montana's bounds")
            
            return Point(long, lat)
        except Exception as e:
            print(f"Error converting coordinates: {str(e)}")
            # Return a point outside Montana's bounds which will be filtered out
            return Point(0, 0)

    def get_shapefile_path(self):
        """Get the path to the shapefile, works for both development and PyInstaller"""
        shapefile_name = 'cb_2021_us_county_5m.shp'
        if getattr(sys, '_MEIPASS', False):
            # If running as compiled executable
            base_path = sys._MEIPASS
        else:
            # If running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        shapefile_path = os.path.join(base_path, 'shapefiles', shapefile_name)
        if not os.path.exists(shapefile_path):
            raise FileNotFoundError(f"Shapefile not found at: {shapefile_path}")
        return shapefile_path

    def generate_map(self):
        """Generate the county-based maps"""
        if self.excel_data is None:
            self.toast.show_toast("Please load an Excel file first", error=True)
            return
        
        try:
            loading = LoadingIndicator(self.root, "Generating county maps...")
            
            # Load the shapefile using the correct path
            try:
                shapefile_path = self.get_shapefile_path()
                counties = gpd.read_file(shapefile_path)
            except Exception as e:
                loading.destroy()
                self.toast.show_toast(f"Error loading shapefile: {str(e)}", error=True)
                return
            
            # Filter for Montana counties
            self.montana_counties = counties[counties['STATEFP'] == '30'].copy()
            
            # Get year input
            try:
                selected_year = int(self.year_var.get())
                if selected_year <= 0:
                    loading.destroy()
                    raise ValueError("Year must be a positive number")
            except ValueError:
                loading.destroy()
                self.toast.show_toast("Please enter a valid year", error=True)
                return
            
            # Validate Excel data
            required_columns = ['lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', 'year']
            if not all(col in self.excel_data.columns for col in required_columns):
                loading.destroy()
                self.toast.show_toast("Excel file must contain 'lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', and 'year' columns", error=True)
                return
            
            # Get species selection
            fam = self.selected_family.get().strip()
            gen = self.selected_genus.get().strip()
            spec = self.selected_species.get().strip()
            
            if not fam or fam == "Select Family" or not gen or gen == "Select Genus" or not spec or spec == "Select Species":
                loading.destroy()
                self.toast.show_toast("Please select Family, Genus, and Species.", error=True)
                return
            
            loading.update_message("Filtering data...")
            
            # Filter data based on species selection
            filtered = self.excel_data.copy()
            
            # Apply family filter (case-insensitive)
            filtered = filtered[filtered['family'].str.lower() == fam.lower()]
            
            # Apply genus filter (case-insensitive)
            filtered = filtered[filtered['genus'].str.lower() == gen.lower()]
            
            # Apply species filter (case-insensitive) only if not 'all'
            if spec.lower() != 'all':
                filtered = filtered[filtered['species'].str.lower() == spec.lower()]
            
            if len(filtered) == 0:
                loading.destroy()
                self.toast.show_toast("No data found for selected species", error=True)
                return
            
            loading.update_message("Converting coordinates...")
            
            # Convert coordinates to points
            geometries = filtered.apply(self.convert_coordinates, axis=1)
            points = gpd.GeoDataFrame(
                filtered,
                geometry=geometries,
                crs="EPSG:4326"
            )
            
            # Filter points within Montana
            montana_boundary = self.montana_counties.dissolve().to_crs("EPSG:4326").geometry.iloc[0]
            points = points[points.geometry.within(montana_boundary)]
            
            if len(points) == 0:
                loading.destroy()
                self.toast.show_toast("No points found within Montana's boundaries", error=True)
                return
            
            # Convert to county CRS
            points = points.to_crs(self.montana_counties.crs)
            
            # Create two datasets based on year
            loading.update_message("Creating year-based datasets...")
            
            # Map A: Data before or equal to selected year
            map_a_data = points[points['year'] <= selected_year].copy()
            
            # Map B: All data
            map_b_data = points.copy()
            
            # Process both maps
            loading.update_message("Processing Map A (≤ selected year)...")
            self.map_a_counties = self.process_county_data(map_a_data)
            
            loading.update_message("Processing Map B (all data)...")
            self.map_b_counties = self.process_county_data(map_b_data)
            
            # Store the year for map titles
            self.selected_year = selected_year
            
            # Display the maps
            loading.update_message("Rendering maps...")
            self.display_maps()
            
            loading.destroy()
            self.toast.show_toast("County maps generated successfully")
            
        except Exception as e:
            if 'loading' in locals():
                loading.destroy()
            self.toast.show_toast(f"Error generating maps: {str(e)}", error=True)
    
    def process_county_data(self, points_data):
        """Process point data and assign colors to counties based on point density"""
        # Create a copy of counties for processing
        counties_with_data = self.montana_counties.copy()
        counties_with_data['point_count'] = 0
        counties_with_data['color'] = 'white'  # Default color for counties with no data
        
        # Count points in each county
        for idx, county in counties_with_data.iterrows():
            points_in_county = points_data[points_data.geometry.within(county.geometry)]
            counties_with_data.at[idx, 'point_count'] = len(points_in_county)
        
        # Assign colors based on ranges
        ranges = []
        for min_var, max_var, color_var in self.color_ranges:
            min_val = float(min_var.get())
            max_val = float('inf') if max_var.get() == "∞" else float(max_var.get())
            ranges.append((min_val, max_val, color_var.get()))
        
        ranges.sort(key=lambda x: x[0])
        
        for min_val, max_val, color in ranges:
            mask = (counties_with_data['point_count'] >= min_val) & (counties_with_data['point_count'] <= max_val)
            counties_with_data.loc[mask, 'color'] = color
        
        return counties_with_data
    
    def display_maps(self):
        """Display both Map A and Map B vertically with legend at bottom right"""
        if self.current_maps is None:
            return
        
        # Clear the figure
        self.figure.clf()
        
        # Create gridspec for layout control - single column, two rows
        gs = self.figure.add_gridspec(2, 1, height_ratios=[1, 1])
        
        # Add taxonomic hierarchy title at the top center
        species_info = self.current_maps['species_info']
        if species_info:
            family, genus, species = species_info.split(' > ')
            title = f"{family} > {genus} > {species}"
            self.figure.suptitle(title, x=0.5, y=0.99, 
                               ha='center', va='top',
                               fontsize=14, fontweight='bold')
        
        # Create subplots for both maps
        self.ax1 = self.figure.add_subplot(gs[0])  # Map A (top)
        self.ax2 = self.figure.add_subplot(gs[1])  # Map B (bottom)
        
        # Configure both axes
        for ax in [self.ax1, self.ax2]:
            ax.set_frame_on(False)  # Remove frame
            ax.set_xticks([])
            ax.set_yticks([])
            ax.set_aspect('equal')
        
        # Plot Map A (≤ selected year)
        map_a_counties = self.current_maps['map_a']
        for idx, county in map_a_counties.iterrows():
            self.ax1.fill(county.geometry.exterior.xy[0], 
                         county.geometry.exterior.xy[1],
                         facecolor=county['color'],
                         edgecolor='black',
                         linewidth=0.5)
        
        # Plot Map B (all data)
        map_b_counties = self.current_maps['map_b']
        for idx, county in map_b_counties.iterrows():
            self.ax2.fill(county.geometry.exterior.xy[0], 
                         county.geometry.exterior.xy[1],
                         facecolor=county['color'],
                         edgecolor='black',
                         linewidth=0.5)
        
        # Set bounds for both maps
        bounds = self.montana_counties.total_bounds
        padding = (bounds[2] - bounds[0]) * 0.05  # Small padding
        
        for ax in [self.ax1, self.ax2]:
            ax.set_xlim([bounds[0] - padding, bounds[2] + padding])
            ax.set_ylim([bounds[1] - padding, bounds[3] + padding])
        
        # Add A and B labels centered at the top of each map
        selected_year = self.current_maps['selected_year']
        self.ax1.text(0.5, 0.98, f'Map A <= {selected_year}', transform=self.ax1.transAxes,
                     fontsize=12, fontweight='bold', va='top', ha='center')
        self.ax2.text(0.5, 0.98, 'Map B: All data', transform=self.ax2.transAxes,
                     fontsize=12, fontweight='bold', va='top', ha='center')
        
        # Create legend elements
        import matplotlib.patches as mpatches
        legend_elements = []
        legend_labels = ['0', '1-10', '11-100', '101-1000', '>1000']
        colors = ['white', '#e7e8e9', '#bcbec0', '#939598', '#231f20']
        
        for color, label in zip(colors, legend_labels):
            legend_elements.append(mpatches.Patch(facecolor=color, 
                                                edgecolor='black',
                                                label=label))
        
        # Add legend to bottom right of Map B
        legend = self.ax2.legend(handles=legend_elements,
                               loc='lower right',
                               frameon=False,
                               bbox_to_anchor=(1.2, 0.165),
                               ncol=1)
        
        # Adjust layout - reduce vertical space between maps
        self.figure.subplots_adjust(left=0.05, right=0.95, 
                                  bottom=0.05, top=0.95,
                                  hspace=0.02)  # Reduced vertical spacing
        
        # Draw the canvas
        self.canvas.draw()

    def download_map(self):
        if self.current_maps is None:
            self.toast.show_toast("Please generate maps first", error=True)
            return
            
        try:
            import datetime
            from pathlib import Path
            import os
            
            # Get Downloads folder path
            downloads_path = str(Path.home() / "Downloads")
            
            # Get current date and time in the desired format
            now = datetime.datetime.now()
            timestamp = now.strftime("%I_%M_%p_%m_%d_%Y")  # e.g., 12_49_PM_6_12_2025
            
            # Create a meaningful filename
            filename = f"MontanaCountyMaps_{timestamp}.tiff"
            file_path = os.path.join(downloads_path, filename)
            
            # Save the figure
            self.figure.savefig(file_path, format="tiff", dpi=300, bbox_inches='tight')
            
            # Show toast notification
            self.toast.show_toast(f"Maps saved as {filename}")
            
            print(f"✅ TIFF maps saved as '{file_path}'")
            
        except Exception as e:
            messagebox.showerror("Error", 
                f"Error saving file:\n{str(e)}\n\n"
                "Please try again."
            )

    def on_window_resize(self, event=None):
        # Update the figure size to match the panel size
        w = self.right_panel.winfo_width() / 100
        h = self.right_panel.winfo_height() / 100
        self.figure.set_size_inches(w, h)
        
        # If we have maps displayed, redraw them to maintain proper layout
        if self.current_maps is not None:
            self.display_maps()
        else:
            self.canvas.draw()

    def update_genus_dropdown(self, event=None):
        family = self.selected_family.get().strip()
        
        if family == "Select Family":
            self.genus_dropdown["values"] = []
            self.genus_dropdown.set("Select Genus")
            return
        
        # Filter based on family selection
        if family == "All":
            # Get all non-empty genus values
            filtered = self.excel_data[self.excel_data["genus"].notna() & (self.excel_data["genus"].str.strip() != "")]
        else:
            # Get genus for specific family (case-insensitive)
            filtered = self.excel_data[self.excel_data["family"].str.lower() == family.lower()]
        
        # Get valid genera (non-empty/non-null values)
        valid_genera = sorted(filtered["genus"].dropna().unique())
        valid_genera = [g for g in valid_genera if str(g).strip() and str(g).lower() != 'nan']  # Remove empty strings and 'nan'
        
        # Create genus list with special options
        genus_values = ["All"] + [g.title() for g in valid_genera]
        
        # Update Genus dropdown
        self.genus_dropdown["values"] = genus_values
        self.genus_dropdown.set("Select Genus")
        
        # Reset species dropdown
        self.species_dropdown.set("Select Species")
        self.species_dropdown["values"] = []
    
    def update_species_dropdown(self, event=None):
        family = self.selected_family.get().strip()
        genus = self.selected_genus.get().strip()
        
        if family == "Select Family" or genus == "Select Genus":
            self.species_dropdown["values"] = []
            self.species_dropdown.set("Select Species")
            return
        
        # Start with base DataFrame
        filtered = self.excel_data
        
        # Apply family filter
        if family == "All":
            filtered = filtered[filtered["family"].notna() & (filtered["family"].str.strip() != "")]
        else:
            filtered = filtered[filtered["family"].str.lower() == family.lower()]
        
        # Apply genus filter
        if genus == "All":
            filtered = filtered[filtered["genus"].notna() & (filtered["genus"].str.strip() != "")]
        else:
            filtered = filtered[filtered["genus"].str.lower() == genus.lower()]
        
        # Get valid species (non-empty/non-null values)
        valid_species = sorted(filtered["species"].dropna().unique())
        valid_species = [s for s in valid_species if str(s).strip() and str(s).lower() != 'nan']  # Remove empty strings and 'nan'
        
        # Create species list with special options - note lowercase for species
        species_values = ["all"] + valid_species
        
        # Update Species dropdown
        self.species_dropdown["values"] = species_values
        self.species_dropdown.set("Select Species")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    os.environ['GDAL_DATA'] = os.path.join(base, 'gdal-data')
    os.environ['PROJ_LIB'] = os.path.join(base, 'proj')
    app = MainApplication()
    app.run() 