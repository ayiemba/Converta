# Converta

Converta is a simple Windows application for converting coordinates in Excel files. It allows users to upload an Excel file, select the columns containing longitude and latitude, and outputs a new Excel file with converted coordinates in new columns.

## Features
- Upload Excel files (.xlsx, .xls)
- Autodetect and select coordinate columns
- Converts coordinates to decimal degrees
- Outputs new columns: `Longitude_Converted` and `Latitude_Converted`
- Date-stamped output file
- Simple UI with a Close button

## Supported Coordinate Formats
The conversion logic in Converta can handle the following cases:

- **Decimal Degrees**: e.g., `45.123N`, `-120.456`, `120.456W`
- **Degrees, Minutes, Seconds (DMS)**: e.g., `45°30'15"N`, `120°30'15"W`
- **Degrees and Minutes**: e.g., `45°30'N`, `120°30'W`
- **Plain numbers**: e.g., `45.123`, `-120.456`

The application will automatically parse and convert these formats to decimal degrees.

## Requirements
- Python 3.8 or newer
- pandas

## Installation
1. Install Python from [python.org](https://www.python.org/downloads/).
2. Install pandas:
   ```sh
   pip install pandas
   ```

## How to Run
1. Double-click the batch file `run_converta.bat` (see below) or run the following command in your terminal:
   ```sh
   python main.py
   ```

## Windows Batch File
Create a file named `run_converta.bat` in the same folder as `main.py` with the following content:

```bat
@echo off
python main.py
```

Double-click `run_converta.bat` to launch the Converta application.

## Usage
1. Click **Browse** to select your Excel file.
2. The app will autodetect and select the most likely longitude and latitude columns. You can change these if needed.
3. Click **Convert** to process the file. The converted file will be saved in the same folder with a date-stamped name.
4. Click **Close** to exit the application.

## Output
- The converted Excel file will have three new columns:
  - `Longitude_Converted`
  - `Latitude_Converted`
  - `Convert_Status`: Shows `Yes` if both coordinates were successfully converted, `No` if either could not be converted.
- The filename will include a date stamp, e.g., `yourfile_converted_20250811.xlsx`.

## License
MIT
