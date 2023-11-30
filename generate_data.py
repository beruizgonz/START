import pandas as pd
import openpyxl
import os
import random as random

# Define general variables
PERIODS = 5
N_PERSONS = 3
PERIODS_EMERGENCY = 10

# Create an save an excel file
excel_file_path = os.path.join(os.getcwd(), 'Simulate_data.xlsx')
def create_excel_file(excel_file_path):
    wb = openpyxl.Workbook()
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
    # Create the data frame   
    df = pd.DataFrame(index=persons, columns=periods)
    # Fill the dataframe with random values (0, 0.5, 1) that represent the availability
    values = [0, 0.5, 1]
    for i in range(N_PERSONS):
        for j in range(PERIODS):
            df.iloc[i, j] = random.choice(values)
    # Add the excel sheet 
    add_sheet_excel(excel_file_path, 'Availability', df)

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
            abbreviation += word[0]
        abbreviation = abbreviation.upper()
        dict_abbreviations[key] = abbreviation
    df_profiles = df_profiles.rename(columns={0: 'Profile', 1: 'Number of people needed'})
    df_profiles = df_profiles.drop(0)
    df_profiles['Code'] = df_profiles.index
    df_profiles['Abbreviation'] = df_profiles[2].map(dict_abbreviations)
    # Change the order of the columns
    df_profiles = df_profiles[['Code', 'Abbreviation', 'Profile', 'Number of people needed']]
    add_sheet_excel(excel_file_path, 'Demand', df_profiles, False )

def create_periods_emergency(periods_emergency = PERIODS_EMERGENCY):
    # Create a list with the periods_emergency
    periods_e = []
    for i in range(periods_emergency):
        periods_e.append("Period " + str(i))
    # Create a dataframe with periods_emergency rows and 2 colums (Outward and Return)
    df_prices = pd.DataFrame(index=periods_e, columns=['Outward', 'Return', 'Emergency'])
    print(df_prices.head())
    # Fill the dataframe with random values (100, 1000)
    for i in range(PERIODS_EMERGENCY):
        df_prices.loc[periods_e[i], 'Outward'] = random.randint(100, 1000)
        df_prices.loc[periods_e[i], 'Return'] = random.randint(100, 1000)

    # Set a constant value (10000000) for the 'Emergency' column
    df_prices['Emergency'] = 10000000
    # Add the excel sheet
    add_sheet_excel(excel_file_path, 'Prices', df_prices)

if __name__ == '__main__':
    create_excel_file(excel_file_path)
    # Create the sheet 'Availability' in the excel file
    persons_availability(PERIODS, N_PERSONS)
    # Create the sheet 'Demand' in the excel file
    personal_START_team()
    # Create the sheet 'Prices' in the excel file
    create_periods_emergency()




