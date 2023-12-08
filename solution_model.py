import os
import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import numpy as np
import openpyxl

from utils import create_color_palettes, export_to_excel
from generate_data import add_sheet_excel

# LECTURA DE DATOS
datafile='Data(1).xlsx'

shdata = pd.read_excel(datafile, sheet_name='Data', header=None)
pNperiods = shdata.loc[0,1]
pNhealthp = shdata.loc[1,1]
pNpeople = shdata.loc[2,1]
pNcharter = shdata.loc[3,1]
pDiscount = shdata.loc[4,1]
pKpeople = shdata.loc[5,1]
pMaxperiods = shdata.loc[6,1]
pMinperiods = shdata.loc[7,1]

shprices = pd.read_excel(datafile, sheet_name='Prices')
pCostOut = shprices.loc[:,'Outward'].to_numpy()
pCostRet = shprices.loc[:,'Return'].to_numpy()

shcharter = pd.read_excel(datafile, sheet_name='Charter')
pCostChar = shcharter.loc[0:pNperiods-1,'Chartered 1':'Chartered 2'].to_numpy()
pMinCapCh = shcharter.loc[0:pNcharter-1,'Min Cap'].to_numpy()
pMaxCapCh = shcharter.loc[0:pNcharter-1,'Max Cap'].to_numpy()

shdemand = pd.read_excel(datafile, sheet_name='Demand')
pHealthdem = shdemand.loc[:,1:].to_numpy()

shhealthprofiles = pd.read_excel(datafile, sheet_name='HealthProfiles', header=None)
pNameprofiles = shhealthprofiles.loc[0,1:].to_numpy()
pNameabbprofiles = shhealthprofiles.loc[1,1:].to_numpy()
pProfiles = shhealthprofiles.loc[2:,1:].to_numpy()

shavailability = pd.read_excel(datafile, sheet_name='Availability')
pAvailability = shavailability.loc[:,1:].to_numpy()

shweights = pd.read_excel(datafile, sheet_name='Weights', header=None)
pW1 = shweights.loc[0,1]
pW2 = shweights.loc[1,1]
pW3 = shweights.loc[2,1]

pWW1 = shweights.loc[4,1]
pWW2 = shweights.loc[5,1]
pWW3 = shweights.loc[6,1]
pWW4 = shweights.loc[7,1]

def model_dimensions(model, excel):
    # Create a dataframe with the 5 rows and 2 columns
    df = pd.DataFrame(index = ['total', 'continuous', 'binary', 'integer', 'constrains'], columns=['Type', 'Number'])
    df['Type'] = df.index
    # Assuming you have 'model' as the optimization model, you can replace these values
    df['Number'] = [    
        model.NumVars,
        model.NumVars-model.NumIntVars,
        model.NumBinVars,
        model.NumIntVars-model.NumBinVars, 
        model.NumConstrs
    ]
    # df.reset_index(drop=True, inplace=True)
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Model dimensions', df, index = True)
# Add later
# Delete column from the excel file
# sheet.delete_cols(1)
# wb.save(excel_file_path)

def model_performance(model,excel): 
    info = ['Objective value', 'Runtime', 'MIPGap', 'Status']
    df = pd.DataFrame(index = info, columns=['Value'])
    # Assuming you have 'model' as the optimization model, you can replace these values
    df['Value'] = [
        model.objVal,
        model.Runtime,
        model.MIPGap,
        model.Status
    ]
    # df.reset_index(drop=True, inplace=True)
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Model performance', df, index = True)
    #df.to_excel(excel, sheet_name='Performance')

def chartered(model, excel): 
    # Access vGamma variables
    vGamma = {}
    for var in model.getVars():
        if 'Gamma' in var.VarName:
            i, t = map(int, var.VarName.split('_')[1:])
            vGamma[i, t] = var
    charter_data = []
    cost_data = []
    # Convert vGamma to a NumPy array for efficient indexing
    vGamma_array = np.array([[vGamma[t, l].X for l in range(pNcharter)] for t in range(pNperiods)])
    
    for l in range(pNcharter):
        nl = np.sum(vGamma_array[:, l])
        costl = np.sum(pCostChar[:, l] * vGamma_array[:, l])
    # Create a dataframe with index the type of charted and the columns the cost
        charter_data.append(nl)
        cost_data.append(costl)
    df = pd.DataFrame(index = ['Chartered 1', 'Chartered 2'], columns=['Number', 'Cost'])
    df['Number'] = charter_data
    df['Cost'] = cost_data
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Charter', df, index = True)
    #df.to_excel(excel, sheet_name='Chartered')

def regular_and_cost(excel, model): 
    costoutflights = 0
    costretflights = 0
    # Access to the varialbes vXstand, vXdisc, vYstand, vYdisc
    vXstand = {}
    vXdisc = {}
    vYstand = {}
    vYdisc = {}
    for var in model.getVars():
        if 'Xstand' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vXstand[t] = var
        if 'Xdisc' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vXdisc[t] = var
        if 'Ystand' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vYstand[t] = var
        if 'Ydisc' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vYdisc[t] = var
    for t in range(pNperiods - 1):
        costoutflights += pCostOut[t] * vXstand[t].X + (pCostOut[t] - pDiscount) * vXdisc[t].X
        costretflights += pCostRet[t + 1] * vYstand[t + 1].X + (pCostRet[t + 1] - pDiscount) * vYdisc[t + 1].X

    totcostreg = costoutflights + costretflights
    # Add this information to the excel
    df = pd.DataFrame(index = ['Outward', 'Return','Total'], columns=['Value'])
    df['Value'] = [costoutflights, costretflights, totcostreg]
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Regular', df, index = True)

def out_and_return(excel, model):
    # Get the variables vXstand, vXdisc, vYstand, vYdisc, vZout, vZret
    vXstand = {}
    vXdisc = {}
    vYstand = {}
    vYdisc = {}
    vZout = {}
    vZret = {}
    for var in model.getVars():
        if 'Xstand' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vXstand[t] = var
        if 'Xdisc' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vXdisc[t] = var
        if 'Ystand' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vYstand[t] = var
        if 'Ydisc' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vYdisc[t] = var
        if 'Zout' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vZout[t] = var
        if 'Zret' in var.VarName:
            t = int(var.VarName.split('_')[1])
            vZret[t] = var

    npeopstand = 0
    npeopdisc = 0
    npeopchar = 0
    for t in range(0, pNperiods-1):
        npeopstand = npeopstand + vXstand[t].X
        npeopdisc = npeopdisc + vXdisc[t].X
        npeopchar = npeopchar + vZout[t].X

    npeopstandr = 0
    npeopdiscr = 0
    npeopcharr = 0
    for t in range(1, pNperiods):
        npeopstandr += vYstand[t].X
        npeopdiscr += vYdisc[t].X
        npeopchar += vZret[t].X

    # Create a dataframe with the number of people travelling outwards
    df = pd.DataFrame(index = ['Standar', 'Discounted', 'Chartered'], columns=['Number of people outwards', 'Number of people return'])
    df['Number of people outwards'] = [npeopstand, npeopdisc, npeopchar]
    df['Number of people return'] = [npeopstandr, npeopdiscr, npeopcharr]
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Outwards and return', df, index = True)
 
def fligths_plan(excel, model): 
    # Acces to the variables vAlphaOut, vAlphaRet
    vAlphaOut = {}
    vAlphaRet = {}
    for var in model.getVars():
        if 'Alphaout' in var.VarName:
            i, t = map(int, var.VarName.split('_')[1:])
            vAlphaOut[i,t] = var
        if 'Alpharet' in var.VarName:
            i, t = map(int, var.VarName.split('_')[1:])
            vAlphaRet[i,t] = var
    # Convert vAlphaOut and vAlphaRet to a NumPy array for efficient indexing
    dict = {}
    perdep = 0
    perret = 0
    for i in range(0, pNpeople):
        for t in range(0, pNperiods):
            if (vAlphaOut[i,t].X == 1):
                perdep = t
            if (vAlphaRet[i,t].X == 1):
                perret = t
        dict[f'Person {i}'] = [perdep+1, perret+1]
    df_people = pd.DataFrame.from_dict(dict, orient='index', columns=['Departure', 'Return'])
    # Get the dict in the data frame. The index is the name of the person and the columns the departure and return period
    df_people = pd.DataFrame.from_dict(dict, orient='index', columns=['Departure', 'Return'])
    # Add a column with the name of the person
    df_people['Person'] = df_people.index
    # Reorder the columns
    df_people = df_people[['Person', 'Departure', 'Return']]    
    # Add the dataframe to the excel file
    add_sheet_excel(excel, 'Flights plan', df_people, index = False)

def health_profiles(model, excel_file):
    results = np.zeros((pNpeople, pNperiods), dtype=int)
    # Acces to the variables vBeta
    vBeta = {}
    for var in model.getVars():
        if 'vBeta' in var.VarName:
            i, j, t = map(int, var.VarName.split('_')[1:])
            vBeta[i, j, t] = var
    # Convert vBeta to a NumPy array for efficient indexing

    for i in range(0,pNpeople):
        for t in range(0,pNperiods):
            results[i,t] = 0
            for j in range (0,pNhealthp):
                if vBeta[i,j,t].X == 1:
                    results[i,t] = j+1
    # Read the reference data of health profiles
    shdata = pd.read_excel('Data(1).xlsx', sheet_name='HealthProfiles', header=None)
    # Read the second row of the data frame and create a dictionary with the health profiles
    dict_health_profiles = dict(zip(range(1, pNhealthp + 1), shdata.loc[1,1:]))
    df = pd.DataFrame(results)
    color_palettes = create_color_palettes(pNhealthp)
    export_to_excel(excel_file, df,dict_health_profiles, color_palettes)

def necessary_people(excel_file): 

    # Access to the variables vUmas
    vUmas = {}
    for var in model.getVars():
        if 'Umas' in var.VarName:
            j, t = map(int, var.VarName.split('_')[1:])
            vUmas[j,t] = var
    # Create a dataframe with the number of people needed for each period (columns) and each health profile (rows)
    df = pd.DataFrame(index = range(0, pNhealthp), columns=range(0, pNperiods))
    shdata = pd.read_excel('Data(1).xlsx', sheet_name='HealthProfiles', header=None)
    # Read the second row of the data frame and create a dictionary with the health profiles
    dict_health_profiles = dict(zip(range(1, pNhealthp + 1), shdata.loc[1,1:]))
    for j in range(0, pNhealthp):
        for t in range(0, pNperiods):
            df.loc[j, t] = vUmas[j, t].X
    # Add the dataframe to the excel file
    df.index = dict_health_profiles.values()
    for i in range(0, pNperiods):
        df.rename(columns={i: 'Period ' + str(i+1)}, inplace=True)
    add_sheet_excel(excel_file, 'Necessary people', df, index = True)

if __name__ == '__main__': 
    model_path = os.path.join(os.getcwd(), 'model.mps')
    # # Read the model file
    model = gp.read(model_path)
    # Now the new_model should have the same variables and types as the original model

    # Solve the new model
    model.optimize()
    filename = f'Res_{datafile}'
    excel_file_path = os.path.join(os.getcwd(), filename)
    wb = openpyxl.Workbook()
    wb.active.title = 'Model dimensions'
    wb.save(excel_file_path)
    model_performance(model, excel_file_path)
    model_dimensions(model, excel_file_path)
    chartered(model, excel_file_path)
    regular_and_cost(excel_file_path, model)
    out_and_return(excel_file_path, model)
    fligths_plan(excel_file_path, model)
    health_profiles(model, excel_file_path)   
    necessary_people(excel_file_path)
