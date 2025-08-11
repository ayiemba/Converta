# Converta

Converta is a user-friendly Windows application for converting geographic coordinates in Excel files. It helps you quickly transform various coordinate formats including files with mixed formats into decimal degrees, making your data ready for mapping, analysis, or sharing.

## What Converta Does
- **Upload** your Excel file (.xlsx, .xls)
- **Autodetect** longitude and latitude columns (you can adjust if needed)
- **Convert** coordinates to decimal degrees
- **Output** a new Excel file with:
  - `Longitude_Converted`
  - `Latitude_Converted`
  - `Convert_Status` ("Yes" if both coordinates converted, "No" otherwise)
- **Date-stamped** output filename for easy tracking
- **Modern UI** with clear instructions and a Close button

## Supported Coordinate Formats
Converta automatically handles:
- Decimal Degrees: `45.123N`, `-120.456`, `120.456W`
- Degrees, Minutes, Seconds (DMS): `45°30'15"N`, `120°30'15"W`
- Degrees and Minutes: `45°30'N`, `120°30'W`
- Plain numbers: `45.123`, `-120.456`

Rows that cannot be converted are marked with `Convert_Status = No` in the output.

## Requirements
- Python 3.8 or newer
- pandas

## Setup
1. [Download Python](https://www.python.org/downloads/)
2. Open a command prompt and run:
   ```sh
   pip install pandas
   ```

## How to Use Converta
1. Double-click `converta.bat` (included) or run:
   ```sh
   python main.py
   ```
2. In the app:
   - Click **Browse** to select your Excel file
   - Confirm or adjust the detected coordinate columns
   - Review the supported formats in the UI
   - Click **Convert**
   - Find your converted file in the same folder, with a date-stamped name
   - Click **Close** to exit

## Output Example
Your converted Excel file will include:
- `Longitude_Converted` and `Latitude_Converted` columns (decimal degrees)
- `Convert_Status` column ("Yes" or "No")
- Filename like `yourfile_converted_20250811.xlsx`

## License
MIT
