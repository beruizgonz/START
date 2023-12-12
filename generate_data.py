import pandas as pd
import openpyxl
import os
import random as random
from openpyxl.utils import get_column_letter

from utils import add_sheet_excel, columns_dimensions
from opts import parser_args

excel_file_path = os.path.join(os.getcwd(), 'Simulate_data.xlsx')
def create_excel_file(opts):
    print(opts)
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

    for i in range(1, opts.nCharter + 1):
        price_col = f'Chartered {i}'
        df[price_col] = [random.uniform(60000,200000) for _ in range(opts.nPeriods)]

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

def persons_availability_two_years(n_periods, n_persons, excel_file_path):
    periods = ["Period " + str(i) for i in range(n_periods)]
    days = pd.date_range(start="1/1/2023", periods=n_periods, freq='4D', normalize=False)
    persons = ["Person " + str(i) for i in range(n_persons)]
    
    df = pd.DataFrame(index=persons, columns=days)
    values = [0, 1, 2]
    
    for i in range(n_persons):
        for j in range(0, n_periods, 4):
            value = random.choice(values) 
            df.iloc[i, j:j+4] = value
    df.columns = df.columns.strftime('%d/%m/%Y')
    # Make the columns with 10pt widt
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Availability', index=True, startrow=1, header=True)

        # Set the width of the columns
        sheet = writer.sheets['Availability']
        for column in sheet.columns:
            max_length = 0
            column = [column[0]] + [str(cell.value) for cell in column[1:]]
            try:
                max_length = max(len(cell) for cell in column)
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column[0].column_letter].width = adjusted_width
            except (TypeError, AttributeError):
                pass
    # add_sheet_excel(excel_file_path, 'Availability', df, True)
    # wb = openpyxl.load_workbook(excel_file_path)
    # sheet = wb['Availability']
    # columns_dimensions(excel_file_path, wb, sheet, df, width = 12)

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

def health_profiles(n_people, n_profiles, excel_file_path):
    profiles = ["Profile " + str(i+1) for i in range(n_profiles)]
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
    # # Create the sheet 'HealthProfiles' in the excel file
    # # health_profiles(opt)
    persons_availability_two_years(181, opt.nPeople,'Availability2.xlsx')




