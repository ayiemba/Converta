import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import re

def dms_to_decimal(degree, minute, second, direction):
    """
    Converts coordinates from Degrees, Minutes, Seconds (DMS) format to Decimal Degrees (DD) format.

    Args:
        degree (str or float): Degrees component of the coordinate.
        minute (str or float): Minutes component of the coordinate.
        second (str or float): Seconds component of the coordinate.
        direction (str): Cardinal direction ('N', 'S', 'E', 'W').

    Returns:
        float: The coordinate in Decimal Degrees format, rounded to 7 decimal places.
    """
    decimal = float(degree) + float(minute) / 60 + float(second) / 3600
    if direction in ['S', 'W']:
        decimal *= -1
    return round(decimal, 7)

def convert_coordinates(df, x_col, y_col):
    # Dummy conversion: add 1 to each coordinate (replace with your logic)
    df['Converted_X'] = df[x_col] + 1
    df['Converted_Y'] = df[y_col] + 1
    return df

def process_file(filepath, x_col, y_col):
    df = pd.read_excel(filepath)
    df = convert_coordinates(df, x_col, y_col)
    out_path = os.path.splitext(filepath)[0] + '_converted.xlsx'
    df.to_excel(out_path, index=False)
    return out_path

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

def run_conversion():
    file_path = entry_file.get()
    x_col = entry_x.get()
    y_col = entry_y.get()
    if not (file_path and x_col and y_col):
        messagebox.showerror('Error', 'Please provide all inputs.')
        return
    try:
        out_path = process_file(file_path, x_col, y_col)
        messagebox.showinfo('Success', f'Converted file saved as:\n{out_path}')
    except Exception as e:
        messagebox.showerror('Error', str(e))

root = tk.Tk()
root.title('Excel Coordinate Converter')

tk.Label(root, text='Excel File:').grid(row=0, column=0, sticky='e')
entry_file = tk.Entry(root, width=40)
entry_file.grid(row=0, column=1)
tk.Button(root, text='Browse', command=select_file).grid(row=0, column=2)

tk.Label(root, text='X Coordinate Column:').grid(row=1, column=0, sticky='e')
entry_x = tk.Entry(root)
entry_x.grid(row=1, column=1)

tk.Label(root, text='Y Coordinate Column:').grid(row=2, column=0, sticky='e')
entry_y = tk.Entry(root)
entry_y.grid(row=2, column=1)

tk.Button(root, text='Convert', command=run_conversion).grid(row=3, column=1)

root.mainloop()
def parse_coordinate(coord):
    """
    Parses a coordinate string and converts it to Decimal Degrees (DD) format.

    Supports various formats including:
    - Decimal Degrees with optional cardinal direction (e.g., "45.123N", "-120.456").
    - Degrees, Minutes, Seconds (DMS) with cardinal direction (e.g., "45°30'15\"N").
    - Degrees and Minutes with cardinal direction (e.g., "45°30'N").

    Args:
        coord (str): The coordinate string to parse.

    Returns:
        float or None: The coordinate in Decimal Degrees format, or None if parsing fails.
    """
    if pd.isnull(coord):
        return None

    coord = str(coord)
    coord = coord.replace('’', "'").replace('″', '"').replace('“', '"').replace("''", '"')
    coord = coord.strip()

    try:
        # Handle Decimal Degrees format
        if re.match(r'^-?\d+(\.\d+)?[NSEW]?$', coord):
            if coord[-1] in ['N', 'S', 'E', 'W']:
                val = float(coord[:-1])
                if coord[-1] in ['S', 'W']:
                    val *= -1
                return round(val, 7)
            else:
                return round(float(coord), 7)
    except ValueError:
        pass

    # Handle Degrees, Minutes, Seconds (DMS) format
    match = re.match(
        r'(?P<deg>\d+)[°:]?\s*(?P<min>\d+)[\':]?\s*(?P<sec>\d+(?:\.\d+)?)[\"’]?\s*(?P<dir>[NSEW])',
        coord, re.IGNORECASE)
    if match:
        parts = match.groupdict()
        return dms_to_decimal(parts['deg'], parts['min'], parts['sec'], parts['dir'].upper())

    # Handle Degrees and Minutes format
    match = re.match(
        r'(?P<deg>\d+)[°:]?\s*(?P<min>\d+)[\'’]?\s*(?P<dir>[NSEW])',
        coord, re.IGNORECASE)
    if match:
        parts = match.groupdict()
        return dms_to_decimal(parts['deg'], parts['min'], 0, parts['dir'].upper())

    return None

# === MAIN PROCESS ===
def convert_excel_coordinates(input_file, output_file, lat_column='Latitude', lon_column='Longitude'):
    """
    Reads an Excel file, converts latitude and longitude coordinates to Decimal Degrees (DD),
    and saves the results to a new Excel file.

    Args:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to save the output Excel file.
        lat_column (str): Name of the column containing latitude coordinates. Default is 'Latitude'.
        lon_column (str): Name of the column containing longitude coordinates. Default is 'Longitude'.

    Returns:
        None
    """
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Apply conversion
    df['Latitude_DD'] = df[lat_column].apply(parse_coordinate)
    df['Longitude_DD'] = df[lon_column].apply(parse_coordinate)

    # Save to new Excel file
    df.to_excel(output_file, index=False)
    print(f"Converted coordinates saved to: {output_file}")

# === USAGE ===
if __name__ == "__main__":
    # Replace with your file name
    input_excel = "activities.xlsx"
    output_excel = "coordinates_converted.xlsx"
    convert_excel_coordinates(input_excel, output_excel)
