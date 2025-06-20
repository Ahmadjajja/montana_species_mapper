# MontanaSpecimensMapper

A Python application for mapping specimen data across Montana counties using geographic information systems (GIS) and data visualization techniques.

## Overview

MontanaSpecimensMapper is a desktop application that allows users to:
- Load Excel files containing specimen data with geographic coordinates
- Filter data by taxonomic hierarchy (Family, Genus, Species)
- Generate county-based heatmaps showing specimen distribution
- Create comparison maps with year-based filtering
- Export high-resolution maps in TIFF format

## Features

### Data Import
- Supports Excel (.xlsx) files with required columns:
  - `lat`: Latitude coordinates
  - `lat_dir`: Latitude direction (N/S)
  - `long`: Longitude coordinates  
  - `long_dir`: Longitude direction (E/W)
  - `family`: Taxonomic family
  - `genus`: Taxonomic genus
  - `species`: Taxonomic species
  - `year`: Collection year

### Interactive Mapping
- **Species Selection**: Hierarchical dropdown menus for Family → Genus → Species
- **Year Filtering**: Filter data by specific years for temporal analysis
- **Dual Map Display**: 
  - Map A: Shows data up to selected year
  - Map B: Shows all available data
- **Color-coded Counties**: Grayscale intensity based on specimen count ranges

### Export Capabilities
- High-resolution TIFF export (300 DPI)
- Automatic file naming with timestamps
- Saved to user's Downloads folder

## Installation

### Prerequisites
- Python 3.8 or higher
- Windows 10/11 (recommended)

### Setup
1. Clone the repository:
```bash
git clone https://github.com/YourUsername/MontanaSpecimensMapper.git
cd MontanaSpecimensMapper
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
venv\Scripts\activate  # On Windows
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python montana_specimens_mapper.py
```

## Usage

### Loading Data
1. Click "Browse" to select your Excel file
2. The application will validate the file format and display a summary
3. Review the data statistics in the summary dialog

### Creating Maps
1. Select Family, Genus, and Species from the dropdown menus
2. Enter a year for filtering (Map A will show data ≤ this year)
3. Click "Generate County Map" to create the visualization
4. Use "Download County Map" to save as TIFF

### Color Ranges
The application uses a grayscale color scheme:
- **White**: 0 specimens
- **Light Gray**: 1-10 specimens  
- **Medium Gray**: 11-100 specimens
- **Dark Gray**: 101-1000 specimens
- **Black**: >1000 specimens

## File Structure

```
MontanaSpecimensMapper/
├── montana_specimens_mapper.py      # Main application
├── montana_specimens_mapper_icon.py # Icon generation utility
├── requirements.txt                 # Python dependencies
├── MontanaSpecimensMapper.spec      # PyInstaller configuration
├── app_icon.ico                    # Application icon
├── app_icon.png                    # Application icon (PNG)
├── shapefiles/                     # Montana county boundaries
│   ├── cb_2021_us_county_5m.shp
│   ├── cb_2021_us_county_5m.dbf
│   ├── cb_2021_us_county_5m.shx
│   └── ...
└── README.md                       # This file
```

## Building Executable

To create a standalone Windows executable:

```bash
pyinstaller MontanaSpecimensMapper.spec
```

The executable will be created in the `dist/` folder.

## Data Requirements

### Excel File Format
Your Excel file must contain the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| `lat` | Latitude coordinate | 45.1234 |
| `lat_dir` | Latitude direction | N |
| `long` | Longitude coordinate | 110.5678 |
| `long_dir` | Longitude direction | W |
| `family` | Taxonomic family | Asteraceae |
| `genus` | Taxonomic genus | Helianthus |
| `species` | Taxonomic species | annuus |
| `year` | Collection year | 2020 |

### Coordinate Systems
- Supports both decimal degrees and DMS (Degrees, Minutes, Seconds) formats
- Automatically handles coordinate direction (N/S, E/W)
- Validates coordinates within Montana's geographic bounds

## Technical Details

### Dependencies
- **pandas**: Data manipulation and analysis
- **geopandas**: Geographic data processing
- **shapely**: Geometric operations
- **matplotlib**: Map visualization
- **numpy**: Numerical computations
- **Pillow**: Image processing
- **pyinstaller**: Executable creation

### Coordinate Processing
- Converts various coordinate formats to decimal degrees
- Handles Unicode and ASCII degree/minute/second symbols
- Applies directional corrections (N/S, E/W)
- Filters points to Montana's geographic boundaries

### Map Generation
- Uses Montana county shapefiles (2021 Census Bureau)
- Projects coordinates to Montana State Plane (EPSG:32100)
- Creates choropleth maps with customizable color ranges
- Supports high-resolution export

## Troubleshooting

### Common Issues

**"Excel file must contain required columns"**
- Ensure your Excel file has all 8 required columns
- Check column names match exactly (case-sensitive)

**"No data found for selected species"**
- Verify your species selection matches the data
- Check for extra spaces or case differences

**"No points found within Montana's boundaries"**
- Verify coordinates are in the correct format
- Check that coordinates fall within Montana (44°N-49°N, 104°W-116°W)

**Application won't start**
- Ensure all dependencies are installed
- Check Python version (3.8+ required)
- Verify shapefiles are present in the shapefiles/ directory

### Performance Tips
- Use smaller Excel files for faster processing
- Close other applications when generating large maps
- Ensure adequate disk space for TIFF exports

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Montana county boundaries from US Census Bureau
- Geographic data processing with GeoPandas
- Map visualization with Matplotlib 