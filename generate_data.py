import pandas as pd
import openpyxl
import os
import random as random

# Define general variables
PERIODS = 5

N_PEOPLE = 250
MAX_PERIODS = 6
MIN_PERIODS = 4 
N_PERIODS_EMERGENCY = 11
N_PROFILES = 23
DISCOUNT = 40
N_PEOPLE_DISCOUNT = 3
N_DAYS = 2*365

# Create an save an excel file
excel_file_path = os.path.join(os.getcwd(), 'Simulate_data.xlsx')
def create_excel_file(excel_file_path):
    wb = openpyxl.Workbook()
    wb.active.title = 'Data'
    sheet = wb['Data']
    # Add the general data
    sheet['A1'] = 'N periods'
    sheet['B1'] = N_PERIODS_EMERGENCY
    sheet['A2'] = 'N health profiles'
    sheet['B2'] = N_PROFILES
    sheet['A3'] = 'N people'
    sheet['B3'] = N_PEOPLE
    sheet['A4'] = 'Discount'
    sheet['B4'] = DISCOUNT
    sheet['A5'] = 'N. People disc.'
    sheet['B5'] = N_PEOPLE_DISCOUNT
    sheet['A6'] = 'Max periods'
    sheet['B6'] = MAX_PERIODS
    sheet['A7'] = 'Min periods'
    sheet['B7'] = MIN_PERIODS
    wb.save(excel_file_path)

def add_sheet_excel(excel_file, name_sheet, df_data, index = False): 
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        df_data.to_excel(writer, sheet_name=name_sheet, index = index)

def persons_availability(n_periods, n_persons):
    periods = []
    for i in range(n_periods):
        periods.append("Period " + str(i))
    persons = []
    for i in range(n_persons):
        persons.append("Person " + str(i))
    df = pd.DataFrame(index=persons, columns=periods)
    # Fill the dataframe with random values (0, 0.5, 1) that represent the availability
    values = [0, 0.5, 1]
    for i in range(n_persons):
        for j in range(n_periods):
            df.iloc[i, j] = random.choice(values)
    add_sheet_excel(excel_file_path, 'Availability', df, True)

def persons_availability_two_years(n_periods, n_persons):
    periods = []
    for i in range(n_periods):
        periods.append("Period " + str(i))
    days = []
    for i in range(n_periods):
        days.append("Day " + str(i))
    persons = []
    for i in range(n_persons):
        persons.append("Person " + str(i))
    df = pd.DataFrame(index=persons, columns=days)
    values = [0, 0.5, 1]
    for i in range(0,N_PEOPLE):
        for j in range(0, N_DAYS, 4):
            value = random.choice(values)
            df.iloc[i, j:j+4] = value
    add_sheet_excel(excel_file_path, 'Availability', df, True)

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

def create_prices(periods_emergency = N_PERIODS_EMERGENCY):
    periods_e = []
    for i in range(periods_emergency):
        periods_e.append("Period " + str(i))
    df_prices = pd.DataFrame(index=periods_e, columns=['Outward', 'Return', 'Charter'])
    for i in range(N_PERIODS_EMERGENCY):
        df_prices.loc[periods_e[i], 'Outward'] = random.randint(100, 1000)
        df_prices.loc[periods_e[i], 'Return'] = random.randint(100, 1000)
    df_prices['Charter'] = 100000
    add_sheet_excel(excel_file_path, 'Prices', df_prices)

def health_profiles(n_people, n_profiles, excel_file_path):
    profiles = ["Profile " + str(i) for i in range(n_profiles)]
    persons = ["Person " + str(i) for i in range(n_people)]
    df = pd.DataFrame(index=persons, columns=profiles)
    for j in range(n_people):
        indices = random.sample(range(n_profiles), 3)
        df.iloc[j, indices] = 1
        for i in range(n_profiles):
            if i not in indices:
                df.iloc[j, i] = 0
    add_sheet_excel(excel_file_path, 'HealthProfiles', df, index=True)

if __name__ == '__main__':
    create_excel_file(excel_file_path)
    # Create the sheet 'Availability' in the excel file
    persons_availability(N_PERIODS_EMERGENCY, N_PEOPLE)
    # Create the sheet 'Demand' in the excel file
    personal_START_team()
    # Create the sheet 'Prices' in the excel file
    create_prices()
    # Create the sheet 'HealthProfiles' in the excel file
    health_profiles(N_PEOPLE, N_PROFILES, excel_file_path=excel_file_path)




