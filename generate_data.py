import pandas as pd
import openpyxl
import os
import random as random
import numpy as np
from openpyxl.utils import get_column_letter

from utils import add_sheet_excel, columns_dimensions
from opts import parser_args

excel_file_path = os.path.join(os.getcwd(), 'Simulate_data.xlsx')
def create_excel_file(opts):
    wb = openpyxl.Workbook()
    wb.active.title = 'Data'
    sheet = wb['Data']
    # Add the general data
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

def create_prices(opts):
    df_prices = pd.DataFrame(index=range(1,opts.nPeriods +1), columns=['Outward', 'Return'])
    random.seed(opts.seed)
    for i in range(opts.nPeriods):
        df_prices.loc[i+1, 'Outward'] = random.randint(100, 1000)
        df_prices.loc[i+1, 'Return'] = random.randint(100, 1000)
    df_prices['Tperiod'] = df_prices.index
    df_prices = df_prices[['Tperiod', 'Outward', 'Return']]  
    add_sheet_excel(excel_file_path, 'Prices', df_prices, False)

def persons_availability(opts):
    periods = []
    for i in range(opts.nPeriods):
        periods.append("Period " + str(i))
    persons = []
    for i in range(opts.nPeople):
        persons.append("Person " + str(i))
    df = pd.DataFrame(index=persons, columns=periods)
    values = [0, 1, 2]
    for i in range(opts.nPeople):
        for j in range(opts.nPeriods):
            df.iloc[i, j] = random.choice(values)
    add_sheet_excel(excel_file_path, 'Availability', df, True)
    wb = openpyxl.load_workbook(excel_file_path)
    columns_dimensions(excel_file_path, wb, 'Availability', df, width = 10)

def chartered_flights(opts):
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

def persons_availability_two_years(n_periods, n_persons):
    periods = ["Period " + str(i) for i in range(n_periods)]
    days = pd.date_range(start="1/1/2023", periods=n_periods, freq='4D', normalize=False)
    persons = ["Person " + str(i) for i in range(n_persons)]
    
    df = pd.DataFrame(index=persons, columns=days)
    values = [0, 0.5, 1]
    
    for i in range(n_persons):
        for j in range(0, n_periods, 4):
            value = random.choice(values) 
            df.iloc[i, j:j+4] = value
    df.columns = df.columns.strftime('%d/%m/%Y')
    # Make the columns with 10pt widt
    add_sheet_excel(excel_file_path, 'Availability', df, True)
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Availability']
    columns_dimensions(excel_file_path, wb, sheet, df, width = 12)

def demand(opts):
    profile_excel = opts.profile_path
    wb = openpyxl.load_workbook(profile_excel)
    sheet = wb['EMT 2']
    df_profiles = pd.DataFrame(sheet.values)
    print(df_profiles.head())
    df = pd.DataFrame(index = range(1,opts.nProfiles + 1), columns= range(1,opts.nPeriods + 1))
    for i in range(opts.nProfiles):
        df.loc[i+1,:] = df_profiles.iloc[i+1, 2]
    df['Profile | Tperiod'] = df.index
    df = df[['Profile | Tperiod'] + list(df.columns[:-1])]
    add_sheet_excel(excel_file_path, 'Demand', df, False)
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Demand']
    columns_dimensions(excel_file_path, wb, sheet, df, width = 14)
    
def personal_START_team(): 
    # Generate the demand data of person needs.
    parent_dir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
    emt2_data = os.path.join(parent_dir, 'profiles_EMT2.xlsx')
    wb = openpyxl.load_workbook(emt2_data)
    sheet = wb['EMT 2']
    df_profiles = pd.DataFrame(sheet.values)
    dict_profiles = {}
    for i in range(1, len(df_profiles)):
        dict_profiles[i] = df_profiles.iloc[i, 0]
    dict_abbreviations = {}
    for key, value in dict_profiles.items():
        words = value.split()
        abbreviation = ''
        for word in words:
            if word[0].isupper():
                abbreviation += word[0]
            if word == 'Pediatra':
                abbreviation = 'PE'
            if word == 'Psiaquiatría':
                abbreviation = 'PS'
        abbreviation = abbreviation.upper()
        dict_abbreviations[key] = abbreviation
    df_profiles = df_profiles.rename(columns={0: 'Profile', 1: 'Number of people needed'})
    df_profiles = df_profiles.drop(0)
    df_profiles['Code'] = df_profiles.index
    df_profiles['Abbreviation'] = df_profiles['Code'].map(dict_abbreviations)
    df_profiles = df_profiles[['Code',  'Profile', 'Abbreviation', 'Number of people needed']]
    add_sheet_excel(excel_file_path, 'Demand', df_profiles, False )
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb['Demand']
    columns_dimensions(excel_file_path, wb, sheet, df_profiles, width = 30)

def weights(opts):
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

import pandas as pd
import numpy as np
import openpyxl

def apply_binomial(row):
    return [np.random.binomial(1, float(p)) for p in row]

import random

def modify_list(input_list, j):
    if len(input_list) <= 1:
        return input_list  # No modification needed for lists of length 1 or less
    
    # Find indices of all '1's except for the first element
    one_indices = [i for i, x in enumerate(input_list[j:], start=1) if x == 1]

    if not one_indices:
        return input_list  # No '1's to modify

    # Randomly select one index to keep as '1'
    keep_index = random.choice(one_indices)

    # Set all other '1's to '0', except the one at keep_index
    for i in one_indices:
        if i != keep_index:
            input_list[i] = 0

    return input_list

def health_profiles(opts):
    # Read the profiles files
    wb = openpyxl.load_workbook(opts.profile_path)
    sheet = wb['EMT 2']
    df_profiles = pd.DataFrame(sheet.values)
    dict_profiles = {}
    sheet_probability = wb['Probabilities']
    df_probability = pd.DataFrame(sheet_probability.values)
    
    # Delete the first column and the first row
    df_probability = df_probability.drop(0)
    df_probability = df_probability.drop(0, axis=1)
    df_profiles = df_profiles.drop(0)
    for i in range(1, len(df_profiles) + 1):
        dict_profiles[i] = df_profiles.iloc[i-1, 3]
    print(dict_profiles)
    # For every profile, create opts.ratio number of people
    df_health_profiles = pd.DataFrame(columns=range(1, opts.nProfiles + 1))
    k = 0
    # Iterate over the rows in df_probability and apply binomial to each
    for i, row in df_probability.iterrows():
        profile = dict_profiles[i]
        profile = int(profile)
        for j in range(opts.ratio*profile):
            binomial_results = apply_binomial(row)
            if i == 1 or i == 2:
                binomial_results = modify_list(binomial_results, i)
            df_health_profiles.loc[k,:] = binomial_results
            k += 1
    # Add health profile to the dataframe
    df_health_profiles['Person | Profile'] = df_health_profiles.index +1
    df_health_profiles = df_health_profiles[['Person | Profile'] + list(df_health_profiles.columns[:-1])]
    add_sheet_excel(opts.output_path, 'HealthProfiles', df_health_profiles, False)

def availability_complementary(opts): 
    for i in range(opts.nPeople): 
        number_preference = random.randint(0,2)
        probabilities = generate_probabilities(number_preference)
        print(f'Number: {number_preference} and probabilities {probabilities}')
        for i in range(opts.nPeriods):
            df = pd.DataFrame(index=range(1,opts.nPeople + 1), columns= range(1,opts.nPeriods + 1))
            for i in range(opts.nPeople):
                for j in range(opts.nPeriods):
                    df.loc[i+1,j+1] = weighted_random_choice(probabilities)
    add_sheet_excel(opts.output_path, 'Availability', df, True)   

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


    


if __name__ == '__main__':
    opt = parser_args()
    # create_excel_file(opt)
    # # Create the sheet 'Availability' in the excel file
    # #persons_availability_two_years(opt)
    # # Create the sheet 'Demand' in the excel file
    # # personal_START_team()
    # demand(opt)
    # # Create the sheet 'Prices' in the excel file 
    # create_prices(opt)
    # # Create the sheet 'Chartered' in the excel file
    # chartered_flights(opt)
    # # Create the sheet 'Weights' in the excel file
    # weights(opt)
    # Create the sheet 'HealthProfiles' in the excel file
    # health_profiles(opt)
    availability_complementary(opt)




