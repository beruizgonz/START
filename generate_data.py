import pandas as pd
import openpyxl
import os
import random as random
import numpy as np

from utils import add_sheet_excel, columns_dimensions, generate_probabilities, weighted_random_choice
from opts import parser_args

excel_file_path = os.path.join(os.getcwd(), 'Simulate_data.xlsx')
def create_excel_file(opts, excel_file_path):
    wb = openpyxl.Workbook()
    wb.active.title = 'Data'
    sheet = wb['Data']
    sheet['A1'] = 'N periods'
    sheet['B1'] = opts.nPeriods
    sheet['A2'] = 'N health profiles'
    sheet['B2'] = opts.nProfiles
    sheet['A3'] = 'N people'
    sheet['B3'] = opts.nPeople
    sheet['A4'] = 'Number chartered'
    sheet['B4'] = opts.nCharter
    sheet['A5'] = 'Discount'
    sheet['B5'] = opts.Discount
    sheet['A6'] = 'N people discount'
    sheet['B6'] = opts.nDiscount
    sheet['A7'] = 'Max periods'
    sheet['B7'] = opts.maxPeriods
    sheet['A8'] = 'Min periods'
    sheet['B8'] = opts.minPeriods
    sheet.column_dimensions['A'].width = 18
    wb.save(excel_file_path)

def create_prices(opts, excel_file_path):
    """
    Create the sheet 'Prices' in the excel file.
    """
    df_prices = pd.DataFrame(index=range(1,opts.nPeriods +1), columns=['Outward', 'Return'])
    random.seed(opts.seed)
    for i in range(opts.nPeriods):
        df_prices.loc[i+1, 'Outward'] = random.randint(100, 1000)
        df_prices.loc[i+1, 'Return'] = random.randint(100, 1000)
    df_prices['Tperiod'] = df_prices.index
    df_prices = df_prices[['Tperiod', 'Outward', 'Return']]  
    add_sheet_excel(excel_file_path, 'Prices', df_prices, False)


def chartered_flights(opts, excel_file_path):
    """
    Create the sheet 'Chartered' in the excel file.
    """
    df = pd.DataFrame(index=range(1, opts.nPeriods + 1), columns=[f'Chartered {i + 1}' for i in range(opts.nCharter)])
    random.seed(opts.seed)

    df['Chartered 1'] = 75000
    df ['Chartered 2'] = 180000
    min_cap_values = opts.minCapacity
    max_cap_values = opts.maxCapacity

    df.insert(2, '', '')  # Inserting a blank column at index 2

    df['Min cap'] = min_cap_values[:len(max_cap_values)] + [None] * (opts.nPeriods - len(max_cap_values))
    df['Max cap'] = max_cap_values[:len(max_cap_values)] + [None] * (opts.nPeriods - len(max_cap_values))
    df['Tperiod'] = df.index
    df = df[['Tperiod', 'Chartered 1', 'Chartered 2', '', 'Min cap', 'Max cap']]
    add_sheet_excel(excel_file_path, 'Charter', df, False)
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Charter']
    columns_dimensions(excel_file_path, wb, sheet, df, width = 14)

def demand(opts, excel_file_path):
    profile_excel = opts.profile_path
    wb = openpyxl.load_workbook(profile_excel)
    sheet = wb['EMT 2']
    df_profiles = pd.DataFrame(sheet.values)
    df = pd.DataFrame(index = range(1,opts.nProfiles + 1), columns= range(1,opts.nPeriods + 1))
    for i in range(opts.nProfiles):
        df.loc[i+1,:] = df_profiles.iloc[i+1, 3]
    df['Profile | Tperiod'] = df.index
    df = df[['Profile | Tperiod'] + list(df.columns[:-1])]
    add_sheet_excel(excel_file_path, 'Demand', df, False)
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Demand']
    columns_dimensions(excel_file_path, wb, sheet, df, width = 14)
    
def weights(opts, excel_file_path):
    df = pd.DataFrame(index=range(1, 6), columns=['Weight', 'Value ', 'Description'])
    df.loc[1, 'Weight'] = 'W1'
    df.loc[1, 'Value '] = opts.w1
    df.loc[1, 'Description'] = 'Weight for the cost'
    df.loc[2, 'Weight'] = 'W2'
    df.loc[2, 'Value '] = opts.w2
    df.loc[2, 'Description'] = 'Weight for infeasibility'
    df.loc[3, 'Weight'] = 'W3'
    df.loc[3, 'Value '] = opts.w3
    df.loc[3, 'Description'] = 'Weight for second objective'
    df.loc[5, 'Weight'] = 'ww1'
    df.loc[5, 'Value '] = opts.ww1
    df.loc[5, 'Description'] = 'Weight for chartered flights'
    df.loc[6, 'Weight'] = 'ww2'
    df.loc[6, 'Value '] = opts.ww2
    df.loc[6, 'Description'] = 'Weight for regular flights'
    df.loc[7, 'Weight'] = 'ww3'
    df.loc[7, 'Value '] = opts.ww3
    df.loc[7, 'Description'] = 'Weight for number of roles'
    df.loc[8, 'Weight'] = 'ww4'
    df.loc[8, 'Value '] = opts.ww4
    df.loc[8, 'Description'] = 'Weight for availability'
    add_sheet_excel(excel_file_path, 'Weights', df, False)
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Weights']
    # Delete the first row of the sheet
    sheet.delete_rows(1)
    columns_dimensions(excel_file_path, wb, sheet, df, width = 30)

def modify_list(input_list, j):
    if len(input_list) <= 1:
        return input_list 
    one_indices = [i for i, x in enumerate(input_list[j:], start=1) if x == 1]
    if not one_indices:
        return input_list 
    keep_index = random.choice(one_indices)
    for i in one_indices:
        if i != keep_index:
            input_list[i] = 0
    return input_list

def health_profiles(opts, excel_file_path):
    wb = openpyxl.load_workbook(opts.profile_path, excel_file_path)
    sheet = wb['EMT 2']
    df_profiles = pd.DataFrame(sheet.values)
    dict_profiles = {}
    sheet_probability = wb['Probabilities']
    df_probability = pd.DataFrame(sheet_probability.values)
    df_probability = df_probability.drop(0)
    df_probability = df_probability.drop(0, axis=1)
    df_profiles = df_profiles.drop(0)
    for i in range(1, len(df_profiles) + 1):
        dict_profiles[i] = df_profiles.iloc[i-1, 3]
    df_health_profiles = pd.DataFrame(columns=range(1, opts.nProfiles + 1))
    k = 0
    for i, row in df_probability.iterrows():
        profile = dict_profiles[i]
        profile = int(profile)
        for j in range(opts.ratio*profile):
            binomial_results = [np.random.binomial(1, float(p)) for p in row]
            if i == 1 or i == 2:
                binomial_results = modify_list(binomial_results, i)
            df_health_profiles.loc[k,:] = binomial_results
            k += 1
    df_health_profiles['Person | Profile'] = df_health_profiles.index +1
    df_health_profiles = df_health_profiles[['Person | Profile'] + list(df_health_profiles.columns[:-1])]
    add_sheet_excel(opts.output_path, 'HealthProfiles', df_health_profiles, False)

def availability(opts, excel_file_path):
    df = pd.DataFrame(index=range(1, opts.nPeople + 1), columns=range(1, opts.nPeriods + 1))
    for i in range(opts.nPeople):
        number_preference = random.randint(0, 2)
        probabilities = generate_probabilities(number_preference)
        print(f'Person {i+1}: Number Preference = {number_preference}, Probabilities = {probabilities}')
        for j in range(opts.nPeriods):
            df.loc[i+1, j+1] = weighted_random_choice(probabilities)
    add_sheet_excel(opts.output_path, 'Availability', df, True)


if __name__ == '__main__':
    opt = parser_args()
    create_excel_file(opt,excel_file_path)
    demand(opt,excel_file_path)
    create_prices(opt,excel_file_path)
    chartered_flights(opt,excel_file_path)
    weights(opt,excel_file_path)
    health_profiles(opt,excel_file_path)
    availability(opt,excel_file_path)




