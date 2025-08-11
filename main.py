import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
import re
import datetime

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

    # If already a number, just return as float
    if isinstance(coord, (int, float)):
        return float(coord)

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

def convert_coordinates(df, x_col, y_col):
    lon_converted = df[x_col].apply(parse_coordinate)
    lat_converted = df[y_col].apply(parse_coordinate)
    df['Longitude_Converted'] = lon_converted
    df['Latitude_Converted'] = lat_converted
    df['Convert_Status'] = [
        'Yes' if (not pd.isnull(lon) and not pd.isnull(lat)) else 'No'
        for lon, lat in zip(lon_converted, lat_converted)
    ]
    return df

def process_file(filepath, x_col, y_col):
    df = pd.read_excel(filepath)
    df = convert_coordinates(df, x_col, y_col)
    date_stamp = datetime.datetime.now().strftime('%Y%m%d')
    out_path = os.path.splitext(filepath)[0] + f'_converted_{date_stamp}.xlsx'
    df.to_excel(out_path, index=False)
    return out_path

def detect_coordinate_columns(columns):
    # Ensure all column names are strings
    columns = [str(col) for col in columns]
    # Latitude candidates: prioritize 'lat', fallback to 'y'
    lat_candidates = [col for col in columns if 'lat' in col.lower()]
    if not lat_candidates:
        lat_candidates = [col for col in columns if col.lower() == 'y']
    # Longitude candidates: prioritize 'lon'/'lng', fallback to 'x'
    lon_candidates = [col for col in columns if 'lon' in col.lower() or 'lng' in col.lower()]
    if not lon_candidates:
        lon_candidates = [col for col in columns if col.lower() == 'x']
    # Fallbacks
    lat_default = lat_candidates[0] if lat_candidates else columns[0] if columns else ''
    lon_default = lon_candidates[0] if lon_candidates else columns[1] if len(columns) > 1 else columns[0] if columns else ''
    return lat_default, lon_default

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        try:
            df = pd.read_excel(file_path)
            columns = list(df.columns)
            combo_x['values'] = columns
            combo_y['values'] = columns
            lat_default, lon_default = detect_coordinate_columns(columns)
            combo_x.set(lon_default)
            combo_y.set(lat_default)
        except Exception as e:
            messagebox.showerror('Error', f'Failed to read Excel file: {e}')

def update_column_options(filepath):
    try:
        df = pd.read_excel(filepath)
        columns = df.columns.tolist()
        combo_x['values'] = columns
        combo_y['values'] = columns
        combo_x.current(0)
        combo_y.current(1)
    except Exception as e:
        messagebox.showerror('Error', str(e))

def run_conversion():
    file_path = entry_file.get()
    x_col = combo_x.get()
    y_col = combo_y.get()
    if not (file_path and x_col and y_col):
        messagebox.showerror('Error', 'Please provide all inputs.')
        return
    try:
        out_path = process_file(file_path, x_col, y_col)
        messagebox.showinfo('Success', f'Converted file saved as:\n{out_path}')
    except Exception as e:
        messagebox.showerror('Error', str(e))

def close_app():
    # Release resources if needed (e.g., clear DataFrames, etc.)
    # For pandas DataFrames, you can use del or set to None
    # Example: del df or df = None
    root.destroy()

def center_window(window, width=400, height=180):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

root = tk.Tk()
root.title('Excel Coordinate Converter')
center_window(root, 540, 260)

main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(fill='both', expand=True)

info_text = (
    'Supported coordinate formats:\n'
    '- Decimal Degrees: 45.123N, -120.456, 120.456W\n'
    '- Degrees, Minutes, Seconds (DMS): 45°30\'15"N, 120°30\'15"W\n'
    '- Degrees and Minutes: 45°30\'N, 120°30\'W\n'
    '- Plain numbers: 45.123, -120.456\n'
    'Rows that cannot be converted will be marked in the output.'
)

info_label = tk.Label(main_frame, text=info_text, justify='left', fg='navy')
info_label.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 12))

# Excel file input
file_label = tk.Label(main_frame, text='Excel File:')
file_label.grid(row=1, column=0, sticky='e', pady=8)
entry_file = tk.Entry(main_frame, width=40)
entry_file.grid(row=1, column=1, padx=5)
browse_btn = tk.Button(main_frame, text='Browse', command=select_file)
browse_btn.grid(row=1, column=2, padx=5)

# X coordinate dropdown
x_label = tk.Label(main_frame, text='Longitude Column:')
x_label.grid(row=2, column=0, sticky='e', pady=8)
combo_x = ttk.Combobox(main_frame, state='readonly', width=37)
combo_x.grid(row=2, column=1, padx=5)

# Y coordinate dropdown
y_label = tk.Label(main_frame, text='Latitude Column:')
y_label.grid(row=3, column=0, sticky='e', pady=8)
combo_y = ttk.Combobox(main_frame, state='readonly', width=37)
combo_y.grid(row=3, column=1, padx=5)

# Buttons
convert_btn = tk.Button(main_frame, text='Convert', command=run_conversion, width=15)
convert_btn.grid(row=4, column=1, pady=18, sticky='e')
close_btn = tk.Button(main_frame, text='Close', command=close_app, width=15)
close_btn.grid(row=4, column=2, pady=18, padx=5, sticky='w')

root.mainloop()

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
