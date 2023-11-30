import os
import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import numpy as np
from utils import create_color_palettes, export_to_excel

# Path to the model files
model_path = os.path.join(os.getcwd(), 'model.lp')


# Read the model file
model = gp.read(model_path)

# Access Gurobi variables vBeta
vBeta = {}
for var in model.getVars():
    if 'beta' in var.VarName:
        i, j, t = map(int, var.VarName.split('_')[1:])
        vBeta[i, j, t] = var

# Solve the model   
model.optimize()

print('Objective value: %g' % model.objVal)

# Data reading 
shdata = pd.read_excel('Data(1).xlsx', sheet_name='Data', header=None)
pNperiods = shdata.loc[0, 1]
pNhealthp = shdata.loc[1, 1]
pNpeople = shdata.loc[2, 1]
pDiscount = shdata.loc[3, 1]
pKpeople = shdata.loc[4, 1]
pMaxperiods = shdata.loc[5, 1]

results = np.zeros((pNpeople, pNperiods), dtype=int)
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        results[i, t] = 0
        for j in range(0, pNhealthp):
            if vBeta[i, j, t].x == 1:  # Check if the variable is selected in the optimal solution
                results[i, t] = j + 1

# Read the reference data of health profiles
shdata = pd.read_excel('Data.xlsx', sheet_name='Demand', header=None)
# Read the first two columns and create a reference dictionary
dict_health_profiles = dict(zip(shdata.iloc[1:, 1], shdata.iloc[1:, 0]))
print(dict_health_profiles)
df = pd.DataFrame(results)
color_palettes = create_color_palettes(pNhealthp)
export_to_excel('planning1', df,dict_health_profiles, color_palettes)


