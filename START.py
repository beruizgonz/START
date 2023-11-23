# -*- coding: utf-8 -*-

import pandas as pd
from gurobipy import GRB, Model, quicksum


# LECTURA DE DATOS
shdata = pd.read_excel('Data.xlsx', sheet_name='Data', header=None)
pNperiods = shdata.loc[0,1]
pNhealthp = shdata.loc[1,1]
pNpeople = shdata.loc[2,1]
pDiscount = shdata.loc[3,1]
pKpeople = shdata.loc[4,1]
pMaxperiods = shdata.loc[5,1]

shprices = pd.read_excel('Data.xlsx', sheet_name='Prices')
pCostOut = shprices.loc[:,'Outward'].to_numpy()
pCostRet = shprices.loc[:,'Return'].to_numpy()

shdemand = pd.read_excel('Data.xlsx', sheet_name='Demand')
pHealthdem = shdemand.loc[:,'N. People'].to_numpy()

shhealthprofiles = pd.read_excel('Data.xlsx', sheet_name='HealthProfiles', header=None)
pProfiles = shhealthprofiles.loc[1:,1:].to_numpy()

shavailability = pd.read_excel('Data.xlsx', sheet_name='Availability', header=None)
pAvailability = shavailability.loc[1:,1:].to_numpy()
# Abstract object for the mathematical optimisation model
model = Model("START")

# VARIABLES DEFINITION
# Binary for person i attending with profile j in time period t
vBeta = {}
for i in range(0, pNpeople):
    for j in range(0, pNhealthp):
        for t in range(0, pNperiods):
            vBeta[i, j, t] = model.addVar(ub=pAvailability[i,j], name='beta %s %s %s' % (i, j, t), vtype=GRB.BINARY, )

# Binary for person i taking an outward flight in time period t
vAlphaout = {}
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        vAlphaout[i,t] = model.addVar(name='alphaout %s %s' % (i, t), vtype=GRB.BINARY)

# Binary for person i taking a return flight in time period t
vAlpharet = {}
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        vAlpharet[i,t] = model.addVar(name='alpharet %s %s' % (i, t), vtype=GRB.BINARY)

# Binary auxiliary for applying or not the discount in outward flights
vDeltaout = {}
for t in range(0, pNperiods):
    vDeltaout[t] = model.addVar(name='deltaout %s' % (t), vtype=GRB.BINARY)

# Binary auxiliary for applying or not the discount in return flights
vDeltaret = {}
for t in range(0, pNperiods):
    vDeltaret[t] = model.addVar(name='deltaret %s' % (t), vtype=GRB.BINARY)

# Number of people taking an outward flight
vNout = {}
for t in range(0, pNperiods):
    vNout[t] = model.addVar(name='nout %s' % (t), vtype=GRB.INTEGER)

# Number of people taking a return flight
vNret = {}
for t in range(0, pNperiods):
    vNret[t] = model.addVar(name='nret %s' % (t), vtype=GRB.INTEGER)

# Number of people taking an outward flight with standard fare in time period t
vXstand = {}
for t in range(0, pNperiods):
    vXstand[t] = model.addVar(name='xstand %s' % (t), vtype=GRB.INTEGER)

# Number of people taking an outward flight with discounted fare in time period t
vXdisc = {}
for t in range(0, pNperiods):
    vXdisc[t] = model.addVar(name='xdisc %s' % (t), vtype=GRB.INTEGER)

# Number of people taking a return flight with standard fare in time period t
vYstand = {}
for t in range(0, pNperiods):
    vYstand[t] = model.addVar(name='ystand %s' % (t), vtype=GRB.INTEGER)

# Number of people taking a return flight with discounted fare in time period t
vYdisc = {}
for t in range(0, pNperiods):
    vYdisc[t] = model.addVar(name='ydisc %s' % (t), vtype=GRB.INTEGER)

# OBJECTIVE FUNCTION
# Objective function definition
model.setObjective(
    quicksum(pCostOut[t] * vXstand[t] + (pCostOut[t] - pDiscount) * vXdisc[t] + pCostRet[t] * vYstand[t] + (pCostRet[t] - pDiscount) * vYdisc[t] for t in range(0, pNperiods)),
    sense=GRB.MINIMIZE
)

# CONSTRAINTS DEFINITION
# Health profiles must be covered
for j in range(0, pNhealthp):
    for t in range(0, pNperiods):
        if t < pNperiods - 1:
            model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) >= pHealthdem[j], 'Profiles_%s_%s' %(j,t))

# Maximum number of periods
for i in range(0, pNpeople):
    model.addConstr(quicksum(t*vAlpharet[i,t] for t in range(0, pNperiods)) - quicksum(t*vAlphaout[i,t] for t in range(0,pNperiods)) <= pMaxperiods, 'Maxperiods_%s' %i)

# Controlling the outward and return flights for each person
for i in range(0, pNpeople):
    for t in range (0, pNperiods):
        model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) - quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) == vAlphaout[i,t] - vAlpharet[i,t], 'Vuelos_%s_%s' %(i,t))

# One health profile for person at most for each time period
for i in range(0, pNpeople):
    for t in range(0, pNperiods):
        model.addConstr(quicksum(vBeta[i, j, t] for j in range(0, pNhealthp)) <= 1, 'Oneprofile_%s_%s' % (i, t))

# Each person can fly once at most outward
for i in range(0, pNpeople):
    model.addConstr(quicksum(vAlphaout[i, t] for t in range(0, pNperiods)) <= 1, 'Onceout_%s' % i)

# Each person can fly once at most return
for i in range(0, pNpeople):
    model.addConstr(quicksum(vAlpharet[i, t] for t in range(0, pNperiods)) <= 1, 'Onceret_%s' % i)

# Number of outward are counted
for t in range(0, pNperiods):
    model.addConstr(vNout[t] == quicksum(vAlphaout[i, t] for i in range(0, pNpeople)), 'Countout_%s' % t)

# Number of return flights are counted
for t in range(0, pNperiods):
    model.addConstr(vNret[t] == quicksum(vAlpharet[i, t] for i in range(0, pNpeople)), 'Countret_%s' % t)

# Counting flights with and without discount outward
for t in range(0, pNperiods):
    model.addConstr(vNout[t] == vXstand[t] + vXdisc[t], 'StandDiscout_%s' % t)

# Counting flights with and without discount return
for t in range(0, pNperiods):
    model.addConstr(vNret[t] == vYstand[t] + vYdisc[t], 'StandDiscret_%s' % t)

# Outward standard fare definition
for t in range(0, pNperiods):
    model.addConstr(vXstand[t] <= (pKpeople-1)*vDeltaout[t], 'StandOut_%s' % t)

# Outward discounted fare definition 1
for t in range(0, pNperiods):
    model.addConstr(pKpeople*(1-vDeltaout[t]) <= vXdisc[t], 'DiscOut1_%s' % t)

# Outward discounted fare definition 2
for t in range(0, pNperiods):
    model.addConstr(vXdisc[t] <= pNpeople*(1-vDeltaout[t]), 'DiscOut2_%s' % t)

# Return standard fare definition
for t in range(0, pNperiods):
    model.addConstr(vYstand[t] <= (pKpeople-1)*vDeltaret[t], 'StandRet_%s' % t)

# Return discounted fare definition 1
for t in range(0, pNperiods):
    model.addConstr(pKpeople*(1-vDeltaret[t]) <= vYdisc[t], 'DiscRet1_%s' % t)

# Return discounted fare definition 2
for t in range(0, pNperiods):
    model.addConstr(vYdisc[t] <= pNpeople*(1-vDeltaret[t]), 'DiscRet2_%s' % t)

# Writing model in format lp
model.write('model.lp')

# Optimise the model
model.optimize()

# Display objective function value
#print("Función objetivo:", model.objVal)

