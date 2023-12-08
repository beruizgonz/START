import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
from gurobipy import *
from solution_model import *
# LECTURA DE DATOS
datafile='Data2Inf.xlsx'

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

# Abstract object for the mathematical optimisation model
model = Model("START")

# VARIABLES DEFINITION
# Binary for person i taking an outward flight in time period t
vAlphaout = {}
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        vAlphaout[i,t] = model.addVar(name='vAlphaout_%s_%s' % (i, t), vtype=GRB.BINARY)

# Binary for person i taking a return flight in time period t
vAlpharet = {}
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        vAlpharet[i,t] = model.addVar(name='vAlpharet_%s_%s' % (i, t), vtype=GRB.BINARY)

# Binary for person i attending with profile j in time period t
vBeta = {}
for i in range(0, pNpeople):
    for j in range(0, pNhealthp):
        for t in range(0, pNperiods):
            vBeta[i, j, t] = model.addVar(ub=min(pAvailability[i,t],pProfiles[i,j]), name='vBeta_%s_%s_%s' % (i, j, t), vtype=GRB.BINARY)

# Binary for chartered flight type 1 in period t
vGamma = {}
for t in range(0, pNperiods):
    for l in range(0, pNcharter):
        vGamma[t,l] = model.addVar(name='vGamma_%s_%s' % (t,l), vtype=GRB.BINARY)

# Binary auxiliary for applying or not the discount in outward flights
vDeltaout = {}
for t in range(0, pNperiods):
    vDeltaout[t] = model.addVar(name='vDeltaout_%s' % (t), vtype=GRB.BINARY)

# Binary auxiliary for applying or not the discount in return flights
vDeltaret = {}
for t in range(0, pNperiods):
    vDeltaret[t] = model.addVar(name='vDeltaret_%s' % (t), vtype=GRB.BINARY)

# Binary variable if person i works with role j
vMu = {}
for i in range(0, pNpeople):
    for j in range(0, pNhealthp):
        vMu[i,j] = model.addVar(ub=pProfiles[i,j], name='vMu_%s_%s' %(i,j), vtype=GRB.BINARY)

# Number of people taking an outward flight with standard fare in time period t
vXstand = {}
for t in range(0, pNperiods):
    vXstand[t] = model.addVar(lb=0.0, name='vXstand_%s' % (t), vtype=GRB.INTEGER)

# Number of people taking an outward flight with discounted fare in time period t
vXdisc = {}
for t in range(0, pNperiods):
    vXdisc[t] = model.addVar(lb=0.0, name='vXdisc_%s' % (t), vtype=GRB.INTEGER)

# Number of people taking a return flight with standard fare in time period t
vYstand = {}
for t in range(0, pNperiods):
    vYstand[t] = model.addVar(lb=0.0, name='vYstand_%s' % (t), vtype=GRB.INTEGER)

# Number of people taking a return flight with discounted fare in time period t
vYdisc = {}
for t in range(0, pNperiods):
    vYdisc[t] = model.addVar(lb=0.0, name='vYdisc_%s' % (t), vtype=GRB.INTEGER)

# Number of people taking a chartered flight outwards in period t
vZout = {}
for t in range(0, pNperiods):
    vZout[t] = model.addVar(lb=0.0, name='vZout_%s' % (t), vtype=GRB.INTEGER)

# Number of people taking a chartered flight outwards in period t
vZret = {}
for t in range(0, pNperiods):
    vZret[t] = model.addVar(lb=0.0, name='vZret_%s' % (t), vtype=GRB.INTEGER)

# Slack variable
vUmenos = {}
for j in range(0, pNhealthp):
    for t in range(0, pNperiods):
        vUmenos[j,t] = model.addVar(lb=0.0, name='vUmenos_%s_%s' % (j,t), vtype=GRB.INTEGER)

# Surplus variable
vUmas = {}
for j in range(0, pNhealthp):
    for t in range(0, pNperiods):
        vUmas[j,t] = model.addVar(lb=0.0, name='vUmas_%s_%s' % (j,t), vtype=GRB.INTEGER)


# OBJECTIVE FUNCTION
# Objective function definition treating infeasibility
model.setObjective(
    pW1*(pWW1*quicksum(quicksum(pCostChar[t,l] * vGamma[t,l] for t in range(0,pNperiods)) for l in range(0,pNcharter)) +
         pWW2*(quicksum(pCostOut[t] * vXstand[t] + (pCostOut[t] - pDiscount)* vXdisc[t] for t in range(0,pNperiods-1)) +
         quicksum(pCostRet[t] * vYstand[t] + (pCostRet[t] - pDiscount) * vYdisc[t] for t in range(1,pNperiods)))) +
    pW2*quicksum(quicksum(vUmas[j,t] for j in range(0,pNhealthp)) for t in range(0,pNperiods-1)) +
    pW3*(pWW3*quicksum(quicksum(vMu[i,j] for i in range(0,pNpeople)) for j in range(0,pNhealthp)) +
         pWW4*quicksum(quicksum(quicksum(pAvailability[i,t] * vBeta[i,j,t] for i in range(0, pNpeople)) for j in range(0, pNhealthp)) for t in range(0, pNperiods))),
    sense=GRB.MINIMIZE)

# CONSTRAINTS DEFINITION

# Health profiles must be covered
"""
for j in range(0, pNhealthp):
    for t in range(0, pNperiods-1):
        model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) >= pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))
"""
# Health profiles must be covered treating infeasibility
for j in range(0, pNhealthp):
    for t in range(0, pNperiods-1):
        model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) - vUmenos[j,t] + vUmas[j,t] == pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))

# Minimum number of periods
for i in range(0, pNpeople):
    model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) >= pMinperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Minperiodsalpha_%s' %i)

# Maximum number of periods
for i in range(0, pNpeople):
    model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) <= pMaxperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Maxperiodsalpha_%s' %i)

# Maximum number of periods with beta
for i in range(0, pNpeople):
    model.addConstr(quicksum(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) for t in range(0, pNperiods)) <= pMaxperiods, 'Maxperiodsbeta_%s' %i)

# Controlling the outward and return flights for each person
for i in range(0, pNpeople):
    for t in range (0, pNperiods):
        if t > 0:
            model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) - quicksum(vBeta[i,j,t-1] for j in range(0, pNhealthp)) == vAlphaout[i,t] - vAlpharet[i,t], 'Flights_%s_%s' %(i,t))

# Controlling the outward and return flights for each person in first period
for i in range(0, pNpeople):
    for t in range (0,1):
        model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) == vAlphaout[i,t], 'Flights1_%s_%s' %(i,t))

# One health profile for person at most for each time period
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        model.addConstr(quicksum(vBeta[i, j, t] for j in range(0, pNhealthp)) <= 1, 'Oneprofile_%s_%s' % (i, t))

# Each person can fly once at most outward
for i in range(0, pNpeople):
    model.addConstr(quicksum(vAlphaout[i, t] for t in range(0, pNperiods-1)) <= 1, 'Onceout_%s' % i)

# Each person can fly once at most return
for i in range(0, pNpeople):
    model.addConstr(quicksum(vAlpharet[i, t] for t in range(1, pNperiods)) <= 1, 'Onceret_%s' % i)

# Number of outward are counted
for t in range(0, pNperiods-1):
    model.addConstr(vXstand[t] + vXdisc[t] + vZout[t] == quicksum(vAlphaout[i, t] for i in range(0, pNpeople)), 'Countout_%s' % t)

# Number of return flights are counted
for t in range(1, pNperiods):
    model.addConstr(vYstand[t] + vYdisc[t] + vZret[t] == quicksum(vAlpharet[i, t] for i in range(0, pNpeople)), 'Countret_%s' % t)

# Outward standard fare definition
for t in range(0, pNperiods-1):
    model.addConstr(vXstand[t] <= (pKpeople-1)*vDeltaout[t], 'StandOut_%s' % t)

# Outward discounted fare definition 1
for t in range(0, pNperiods-1):
    model.addConstr(pKpeople*(1-vDeltaout[t]) <= vXdisc[t], 'DiscOut1_%s' % t)

# Outward discounted fare definition 2
for t in range(0, pNperiods-1):
    model.addConstr(vXdisc[t] <= pNpeople*(1-vDeltaout[t]), 'DiscOut2_%s' % t)

# Return standard fare definition
for t in range(1, pNperiods):
    model.addConstr(vYstand[t] <= (pKpeople-1)*vDeltaret[t], 'StandRet_%s' % t)

# Return discounted fare definition 1
for t in range(1, pNperiods):
    model.addConstr(pKpeople*(1-vDeltaret[t]) <= vYdisc[t], 'DiscRet1_%s' % t)

# Return discounted fare definition 2
for t in range(1, pNperiods):
    model.addConstr(vYdisc[t] <= pNpeople*(1-vDeltaret[t]), 'DiscRet2_%s' % t)

# Only one type of chartered flight is hired at most
for t in range(0, pNperiods):
    model.addConstr(quicksum(vGamma[t,l] for l in range(0,pNcharter)) <= 1, 'Onechtype_%s' % t)

# Satisfying minimum capacity of chartered in outward flights
for t in range(0, pNperiods - 1):
    model.addConstr(quicksum(pMinCapCh[l] * vGamma[t,l] for l in range(0,pNcharter)) <= vZout[t], 'Mincapout_%s' % t)

# Satisfying maximum capacity of chartered in outward flights
for t in range(0, pNperiods - 1):
    model.addConstr(vZout[t] <= quicksum(pMaxCapCh[l] * vGamma[t,l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# Satisfying minimum capacity of chartered in return flights
for t in range(1, pNperiods):
    model.addConstr(quicksum(pMinCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)) <= vZret[t], 'Mincapout_%s' % t)

# Satisfying maximum capacity of chartered in return flights
for t in range(1, pNperiods):
    model.addConstr(vZret[t] <= quicksum(pMaxCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# Number of roles that a person plays 1
for i in range(0, pNpeople):
    for j in range(0, pNhealthp):
        model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) <= pMaxperiods * vMu[i,j], 'Roles1_%s_%s' %(i,j))

# Number of roles that a person plays 2
for i in range(0, pNpeople):
    for j in range(0, pNhealthp):
        model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) >= vMu[i,j], 'Roles1_%s_%s' %(i,j))

# Writing model in format lp
model.write('model.lp')
model.write('model.mps')

# Optimise the model
model.optimize()

# Display model dimensions:
print("\nMODEL DIMENSIONS")
print("Number of variables           : {:d}".format(model.NumVars))
print("Number of continuous variables: {:d}".format(model.NumVars-model.NumIntVars))
print("Number of binary variables    : {:d}".format(model.NumBinVars))
print("Number of integer variables   : {:d}".format(model.NumIntVars-model.NumBinVars))
print("Number of constraints         : {:d}".format(model.NumConstrs))

# Display model performance:
print("\nMODEL PERFORMANCE")
print("Objective function value: {0:.2f}".format(model.objVal))
print("Resolution time         : {0:.0f} secs.".format(model.Runtime))
print("Optimality Gap          : {0:.2f}%".format(100*model.MIPGap))


### PRINTING DATA IN SCREEN
# Number of chartered flights and its total cost
print("\nCHARTERED FLIGHTS AND COSTS")

for l in range(0, pNcharter):
    for t in range(0, pNperiods):
        if (vGamma[t,l].X ==1):
            print("There is a chartered flight of type {:d} in period {:d}\n".format(l+1,t+1))

totcostchar = 0
for l in range(0, pNcharter):
    nl = 0
    costl = 0
    for t in range(0, pNperiods):
        nl = nl + vGamma[t,l].X
        costl = costl + pCostChar[t,l] * vGamma[t,l].X
    print("Number of chartered flights of type {:d}: {:.0f}".format(l+1, nl))
    print("Cost for chartered flights of type {:d} : {:.2f}".format(l+1, costl))
    totcostchar = totcostchar + costl

print("\nTotal cost for chartered flights     : {:.2f}".format(totcostchar))

# Total cost for regular flights
print("\nREGULAR FLIGHTS AND COSTS")
costoutflights = 0
costretflights = 0
for t in range(0, pNperiods-1):
    costoutflights = costoutflights + pCostOut[t] * vXstand[t].X + (pCostOut[t] - pDiscount) * vXdisc[t].X

for t in range(1, pNperiods):
    costretflights = costretflights + pCostRet[t] * vYstand[t].X + (pCostRet[t] - pDiscount) * vYdisc[t].X

totcostreg = costoutflights + costretflights
print("Total cost for outward flights: {:.2f}".format(costoutflights))
print("Total cost for return flights : {:.2f}".format(costretflights))
print("Total cost regular flights    : {:.2f}".format(totcostreg))

# Number of people travelling outwards
print("\nNUMBER OF PEOPLE TRAVELLING OUTWARDS")
npeopstand = 0
npeopdisc = 0
npeopchar = 0
for t in range(0, pNperiods-1):
    npeopstand = npeopstand + vXstand[t].X
    npeopdisc = npeopdisc + vXdisc[t].X
    npeopchar = npeopchar + vZout[t].X

print("\nN. people travelling outwards standard fare    : {:.0f}".format(npeopstand))
print("N. people travelling outwards discounted fare  : {:.0f}".format(npeopdisc))
print("N. people travelling outwards in charter flight: {:.0f}".format(npeopchar))

# Number of people travelling return
print("\nNUMBER OF PEOPLE TRAVELLING RETURN")
npeopstand = 0
npeopdisc = 0
npeopchar = 0
for t in range(1, pNperiods):
    npeopstand = npeopstand + vYstand[t].X
    npeopdisc = npeopdisc + vYdisc[t].X
    npeopchar = npeopchar + vZret[t].X

print("\nN. people travelling return standard fare    : {:.0f}".format(npeopstand))
print("N. people travelling return discounted fare  : {:.0f}".format(npeopdisc))
print("N. people travelling return in charter flight: {:.0f}".format(npeopchar))

# When a person departas and returns
print("\nFLIGHTS PLAN")
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        if (vAlphaout[i,t].X == 1):
            perdep = t
        if (vAlpharet[i,t].X == 1):
            perret = t
    print("Person {:d} departs in period {:d} and returns in period {:d}".format(i+1,perdep+1,perret+1))

# Number of roles that a person plays
print("\nPROFILES PLAN")
results = np.zeros((pNpeople,pNperiods),dtype=int)
for i in range (0, pNpeople):
    nroles = 0
    for j in range (0, pNhealthp):
        nroles = nroles + vMu[i,j].X
        for t in range (0, pNperiods):
            if (vBeta[i,j,t].X == 1):
                print("Person {:d} works with profile {} ({}) in period {:d}".format(i+1,pNameprofiles[j],pNameabbprofiles[j],t+1))
    print("Number of roles of person {:d}: {:.0f}".format(i+1, nroles))
# Number of infeasibilities
print("\nNECESSARY PEOPLE")
for j in range(0, pNhealthp):
    for t in range(0, pNperiods):
        if (vUmas[j,t].X > 0.5):
            print("Profile {} needs {:.0f} people in period {:d}".format(pNameprofiles[j], vUmas[j,t].X, t+1))

# EXPORTING RESULTS TO EXCEL
# Creating matrix of results and profiles
results = np.zeros((pNpeople,pNperiods),dtype=int)
for i in range(0,pNpeople):
    for t in range(0,pNperiods):
        results[i,t] = 0
        for j in range (0,pNhealthp):
            if vBeta[i,j,t].X == 1:
                results[i,t] = j+1

# Creating the dataframe
planning = pd.DataFrame(results)
planning.index += 1
planning.columns += 1

# Creating the excel file
#writer = pd.ExcelWriter('Planning.xlsx', engine='xlsxwriter')
#planning.to_excel('Planning.xlsx', sheet_name='Planning')
#print(planning)
planning.to_excel('Planning.xlsx', sheet_name='Planning')


path_to_excel="Planning.xlsx"
example=openpyxl.load_workbook(path_to_excel)
ws = example['Planning'] #Name of the working sheet


col1 = PatternFill(patternType='solid', fgColor='94e0de')
col2 = PatternFill(patternType='solid', fgColor='FC2C03')
col3 = PatternFill(patternType='solid', fgColor='03FCF4')
col4 = PatternFill(patternType='solid', fgColor='35FC03')
col5 = PatternFill(patternType='solid', fgColor='FCBA03')
col6 = PatternFill(patternType='solid', fgColor='FFE333')

for line in ws['B']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

for line in ws['C']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

for line in ws['D']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

for line in ws['E']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

for line in ws['F']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

for line in ws['G']:
    if line.value == 1:
        line.fill = col1
    else:
        if line.value == 2:
            line.fill = col2
        else:
            if line.value == 3:
                line.fill = col3
            else:
                if line.value == 4:
                    line.fill = col4
                else:
                    if line.value == 5:
                        line.fill = col5
                    else:
                        if line.value == 6:
                            line.fill = col6

example.save(path_to_excel)





