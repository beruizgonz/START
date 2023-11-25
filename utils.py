import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os 

N = 10  # Change N to the desired number of color palettes

# The colors are always the same for a given seed. But the set make that later the coolor assigned to a value is random
def create_color_palettes(N, seed = 42):
    color_palettes = set()  # Use a set to ensure uniqueness
  # Set a seed for reproducibility
    np.random.seed(seed)
    while len(color_palettes) < N:
        # Generate a random color palette
        color = "#{:02X}{:02X}{:02X}".format(*np.random.randint(0, 255, size=(3,)))
        color_palettes.add(color)
        
    dict_color = dict(zip(range(1,N+1), color_palettes))
    return dict_color

# Convert hex color to ARGB. Openpyxl uses ARGB format
def hex_to_argb(hex_color):
    hex_color = hex_color.lstrip("#")
    return "FF" + hex_color

def create_random_data_frame(N, pNperiods, pNpeople):
    data = np.random.randint(1, N + 1, size=(pNpeople, pNperiods))  # Adjust range to start from 1
    df = pd.DataFrame(data)
    return df

# Apply color to a cell
def apply_color(value, color_palettes):
    # Get the color from the dictionary
    if value == 0:
        return None
    color_ref = color_palettes[value]
    color = hex_to_argb(color_ref)
    fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
    return fill

def export_to_excel(name,df, dict_references, color_palettes):
    # Create a Workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Solution'
    # Add a row with the column names as 'Period 1', 'Period 2', etc., and a blank cell at the beginning
    for i in range(1, df.shape[1] + 1):
        sheet.cell(row=1, column=i + 1).value = 'Period ' + str(i)
    # Add a column with the row names as 'Person 1', 'Person 2', etc.
    for i in range(1, df.shape[0] + 1):
        sheet.cell(row=i + 1, column=1).value = 'Person ' + str(i)
    # Apply colors to the entire DataFrame and export to Excel
    for i, row in enumerate(df.values):
        for j, cell_value in enumerate(row):
            color_fill = apply_color(cell_value, color_palettes)
            # addt the value to the cell
            if cell_value == 0:
                sheet.cell(row=i + 2, column=j + 2).value = ''
                # sheet.cell(row=i + 1, column=j + 1).fill = None
            else:
                sheet.cell(row=i + 2, column=j + 2).value = '' # If w want to add the profile name to the cell change for dict_references[cell_value]
                sheet.cell(row=i + 2, column=j + 2).fill = color_fill
           
    # add a legend to the sheet with the dict_references
    for i, key in enumerate(dict_references):
        sheet.cell(row=i + 5, column=df.shape[1] + 3).value = dict_references[key]
        color_fill = apply_color(key, color_palettes)
        sheet.cell(row=i + 5, column=df.shape[1] + 2).fill = color_fill

    workbook.save(f'{name}.xlsx')



if __name__ == "__main__":
    N = 6
    color_palettes1 = create_color_palettes(6)
    color_palettes2 = create_color_palettes(6)
    print(color_palettes1)
    print(color_palettes2)
    #df = create_random_data_frame(6, 18,6)
    #export_to_excel(df, color_palettes)




