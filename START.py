import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
from gurobipy import quicksum, Model, GRB
from solution_model import *

#LECTURA DE DATOS
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

# Writing model in format mps
model.write('model.mps')
