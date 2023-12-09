import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


# The colors are always the same for a given seed. But the set make that later the coolor assigned to a value is random
def create_color_palettes(N, seed = 42):
    color_palettes = set()  # Use a set to ensure uniqueness of colors
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

def export_to_excel(excel_file ,df, dict_references, color_palettes):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.create_sheet('Profile plan')
    for i in range(1, df.shape[1]):
        sheet.cell(row=1, column=i + 1).value = 'Period ' + str(i)
    for i in range(1, df.shape[0] + 1):
        sheet.cell(row=i + 1, column=1).value = 'Person ' + str(i)
    for i, row in enumerate(df.values):
        for j, cell_value in enumerate(row):
            color_fill = apply_color(cell_value, color_palettes)
            if cell_value == 0:
                sheet.cell(row=i + 2, column=j + 2).value = ''
            else:
                sheet.cell(row=i + 2, column=j + 2).value = dict_references[cell_value] 
    for i, key in enumerate(dict_references):
        sheet.cell(row=i + 5, column=df.shape[1] + 3).value = dict_references[key]
        color_fill = apply_color(key, color_palettes)
        sheet.cell(row=i + 5, column=df.shape[1] + 2).fill = color_fill
    workbook.save(excel_file)

def access_model_variables(name_variable, index_variables, model): 
    variable = {}
    if index_variables == 1: 
        for var in model.getVars():
            if name_variable in var.VarName:
                i = int(var.VarName.split('_')[1])
                variable[i] = var
    elif index_variables == 2: 
        for var in model.getVars():
            if name_variable in var.VarName:
                i, t = map(int, var.VarName.split('_')[1:])
                variable[i, t] = var
    elif index_variables == 3:
        for var in model.getVars():
            if name_variable in var.VarName:
                i, j, t = map(int, var.VarName.split('_')[1:])
                variable[i, j, t] = var
    return variable

def add_sheet_excel(excel_file, name_sheet, df_data, index = False): 
    # if the excel shate exists overwrite it
    if name_sheet in openpyxl.load_workbook(excel_file).sheetnames:
        wb = openpyxl.load_workbook(excel_file)
        wb.remove(wb[name_sheet])
        wb.save(excel_file)
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        df_data.to_excel(writer, sheet_name=name_sheet, index = index)

def columns_dimensions(excel_file, wb, sheet, df, width = 10):
    for i in range(df.shape[1]+1):
        column_letter = get_column_letter(i+1)
        sheet.column_dimensions[column_letter].width = width
    wb.save(excel_file)


