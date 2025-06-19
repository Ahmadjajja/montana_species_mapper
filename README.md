# MontanaSpeciesMapper

A professional tool for generating species distribution maps across Montana counties. This application creates comparative county-based maps showing species data distribution, with temporal analysis capabilities.

## Building the Executable

To build the executable, follow these steps:

1. Make sure you have Python 3.8 or later installed
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Place your icon file (`app_icon.ico`) in the project root directory
4. Run PyInstaller to create the executable:
   ```
   pyinstaller montana_county_map.spec
   ```

The executable will be created in the `dist` directory.

## Required Files

Before building, ensure you have:
1. `app_icon.ico` - Your application icon file
2. `shapefiles` directory - Contains the Montana county shapefiles
3. All Python source files

## Running the Application

After building, you can run the application by:
1. Double-clicking the executable in the `dist` directory
2. Or running it from the command line:
   ```
   "dist/Montana County Map Generator.exe"
   ```

## Notes

- The application requires the shapefiles directory to be in the same location as the executable
- The icon file will be automatically included in the executable
- All dependencies will be bundled with the executable

---

## Features

- **Species-Based Filtering:**
  - Select by Family, Genus, and Species.
  - Intuitive dropdown interface for taxonomic selection.
- **Year-Based Filtering:**
  - Generates two comparison maps: Map A (data ≤ selected year) and Map B (all data).
  - Allows temporal analysis of species distribution patterns.
- **Point Counting and Coloring:**
  - Each county is colored based on the number of points it contains.
  - Five distinct color ranges for clear data visualization.
  - Only points inside Montana are counted.
- **Interactive GUI:**
  - Modern, resizable interface with left-side controls and right-side live map preview.
  - Year selection for temporal analysis.
  - Toast notifications for user feedback.
- **Export:**
  - Download the generated maps as a high-resolution TIFF file.
  - File is automatically saved to your Downloads folder with a timestamped filename.
- **Robust Data Handling:**
  - Skips and warns about invalid or out-of-bounds coordinates.
  - Handles large datasets efficiently.

---

## Requirements

- Python 3.8+
- See `requirements.txt` for all dependencies:
  - pandas
  - geopandas
  - matplotlib
  - shapely
  - numpy
  - openpyxl
  - pillow

---

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/YourUsername/MontanaSpeciesMapper.git
   cd MontanaSpeciesMapper
   ```
2. (Recommended) Create and activate a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # On Windows
   # or
   source venv/bin/activate  # On Mac/Linux
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Ensure you have the Montana county shapefile in the `shapefiles/` directory.

---

## Usage

1. **Run the application:**
   ```bash
   python montana_species_mapper.py
   ```
2. **Load your Excel file:**
   - Click "Browse" and select your `.xlsx` file.
   - Required columns: `lat`, `lat_dir`, `long`, `long_dir`, `year` (decimal or DMS format supported).
3. **Configure the map:**
   - Enter the year for filtering data (Map A will show data ≤ this year).
   - Select Family, Genus, and Species for filtering.
4. **Generate the county maps:**
   - Click "Generate County Map" to create two maps side by side.
   - Map A shows data before or equal to the selected year.
   - Map B shows all data for comparison.
5. **Export:**
   - Click "Download County Map" to save both maps as a TIFF in your Downloads folder.
   - The filename will include the current date and time.

---

## Input Data Format

Your Excel file must include:
- `lat`: Latitude (decimal or DMS, e.g., `44.695` or `44°41.576'`)
- `lat_dir`: 'N' or 'S'
- `long`: Longitude (decimal or DMS, e.g., `-110.456` or `110°27.360'`)
- `long_dir`: 'E' or 'W'
- `year`: Year of the data point (numeric)
- `family`: Taxonomic family
- `genus`: Taxonomic genus
- `species`: Taxonomic species

Other columns can be present but are not required for mapping.

---

## How It Works

1. **Coordinate Parsing:**
   - All coordinates are parsed and converted to decimal degrees, with direction applied.
2. **Montana Filtering:**
   - Only points inside the actual Montana polygon are kept.
3. **County Assignment:**
   - Each point is assigned to the Montana county it falls within.
4. **Year-Based Filtering:**
   - Data is split into two datasets: before/equal to selected year and all data.
5. **Point Counting:**
   - For each county, the number of Montana points inside is counted for both datasets.
6. **Color Assignment:**
   - Each county is colored according to the defined color ranges.
7. **Dual Map Generation:**
   - Two maps are generated side by side for temporal comparison.
8. **Export:**
   - Both maps can be saved as a single TIFF with a timestamped filename.

---

## Example Output

- Two clean, publication-ready maps of Montana showing county-based species distribution.
- Map A shows historical data (≤ selected year).
- Map B shows complete dataset for comparison.
- Only Montana data is visualized; out-of-state points are ignored.
- Counties with no data remain white.

---

## Shapefile Requirement

- Place the Montana county shapefile (e.g., `cb_2021_us_county_5m.shp` and related files) in a `shapefiles/` directory in your project root.
- The software will automatically extract the Montana counties from this file.

---

## Troubleshooting

- **Invalid coordinates:** The app will warn and skip rows with unparseable or missing coordinates.
- **Points outside Montana:** The app will warn and skip points outside the state polygon.
- **No points in Montana:** If your data has no valid Montana points, the maps will not be generated.
- **Missing year data:** Ensure your Excel file includes a 'year' column with numeric values.
- **Git integration:** See the repo for version control and collaboration.

---

## License

MIT License

---

## Author

[Ahmadjajja/Heat_Map_Generator](https://github.com/Ahmadjajja/Heat_Map_Generator) 