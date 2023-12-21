import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import random 

def create_color_palettes(N, seed = 42):
    color_palettes = set() 
    np.random.seed(seed)
    while len(color_palettes) < N:
        color = "#{:02X}{:02X}{:02X}".format(*np.random.randint(0, 255, size=(3,)))
        color_palettes.add(color)
    dict_color = dict(zip(range(1,N+1), color_palettes))
    return dict_color

def hex_to_argb(hex_color):
    hex_color = hex_color.lstrip("#")
    return "FF" + hex_color

def create_random_data_frame(N, pNperiods, pNpeople):
    data = np.random.randint(1, N + 1, size=(pNpeople, pNperiods))  
    df = pd.DataFrame(data)
    return df

def apply_color(value, color_palettes):
    if value == 0:
        return None
    color_ref = color_palettes[value]
    color = hex_to_argb(color_ref)
    fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
    return fill

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

def weighted_random_choice(probabilities):
    """ Selects a number based on given probabilities. """
    numbers = [0, 1, 2]
    return np.random.choice(numbers, p=probabilities)
   
def generate_probabilities(number_preference):
    highest_probability = random.uniform(0.6, 0.8)
    remaining_probability = 1 - highest_probability
    probabilities = [0, 0, 0]
    second_probability = random.uniform(0, remaining_probability)
    third_probability = remaining_probability - second_probability
    if second_probability > third_probability:
        second_highest_probability, third_highest_probability = second_probability, third_probability
    else: 
        second_highest_probability, third_highest_probability = third_probability, second_probability
    if number_preference == 0:
        probabilities[number_preference] = highest_probability
        probabilities[1] = second_highest_probability
        probabilities[2] = third_highest_probability
    elif number_preference == 1:
        probabilities[number_preference] = highest_probability
        probabilities[0] = second_highest_probability
        probabilities[2] = third_highest_probability
    elif number_preference == 2:
        probabilities[number_preference] = highest_probability
        probabilities[1] = second_highest_probability
        probabilities[0] = third_highest_probability
    return probabilities


