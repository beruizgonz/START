import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
from gurobipy import quicksum, Model
from solution_model import *
# LECTURA DE DATOS
# datafile='Data2Inf.xlsx'

# shdata = pd.read_excel(datafile, sheet_name='Data', header=None)
# pNperiods = shdata.loc[0,1]
# pNhealthp = shdata.loc[1,1]
# pNpeople = shdata.loc[2,1]
# pNcharter = shdata.loc[3,1]
# pDiscount = shdata.loc[4,1]
# pKpeople = shdata.loc[5,1]
# pMaxperiods = shdata.loc[6,1]
# pMinperiods = shdata.loc[7,1]

# shprices = pd.read_excel(datafile, sheet_name='Prices')
# pCostOut = shprices.loc[:,'Outward'].to_numpy()
# pCostRet = shprices.loc[:,'Return'].to_numpy()

# shcharter = pd.read_excel(datafile, sheet_name='Charter')
# pCostChar = shcharter.loc[0:pNperiods-1,'Chartered 1':'Chartered 2'].to_numpy()
# pMinCapCh = shcharter.loc[0:pNcharter-1,'Min Cap'].to_numpy()
# pMaxCapCh = shcharter.loc[0:pNcharter-1,'Max Cap'].to_numpy()

# shdemand = pd.read_excel(datafile, sheet_name='Demand')
# pHealthdem = shdemand.loc[:,1:].to_numpy()

# shhealthprofiles = pd.read_excel(datafile, sheet_name='HealthProfiles', header=None)
# pNameprofiles = shhealthprofiles.loc[0,1:].to_numpy()
# pNameabbprofiles = shhealthprofiles.loc[1,1:].to_numpy()
# pProfiles = shhealthprofiles.loc[2:,1:].to_numpy()

# shavailability = pd.read_excel(datafile, sheet_name='Availability')
# pAvailability = shavailability.loc[:,1:].to_numpy()

# shweights = pd.read_excel(datafile, sheet_name='Weights', header=None)
# pW1 = shweights.loc[0,1]
# pW2 = shweights.loc[1,1]
# pW3 = shweights.loc[2,1]

# pWW1 = shweights.loc[4,1]
# pWW2 = shweights.loc[5,1]
# pWW3 = shweights.loc[6,1]
# pWW4 = shweights.loc[7,1]



# # VARIABLES DEFINITION
# # Binary for person i taking an outward flight in time period t
# vAlphaout = {}
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         vAlphaout[i,t] = model.addVar(name='vAlphaout_%s_%s' % (i, t), vtype=GRB.BINARY)

# # Binary for person i taking a return flight in time period t
# vAlpharet = {}
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         vAlpharet[i,t] = model.addVar(name='vAlpharet_%s_%s' % (i, t), vtype=GRB.BINARY)

# # Binary for person i attending with profile j in time period t
# vBeta = {}
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         for t in range(0, pNperiods):
#             vBeta[i, j, t] = model.addVar(ub=min(pAvailability[i,t],pProfiles[i,j]), name='vBeta_%s_%s_%s' % (i, j, t), vtype=GRB.BINARY)

# # Binary for chartered flight type 1 in period t
# vGamma = {}
# for t in range(0, pNperiods):
#     for l in range(0, pNcharter):
#         vGamma[t,l] = model.addVar(name='vGamma_%s_%s' % (t,l), vtype=GRB.BINARY)

# # Binary auxiliary for applying or not the discount in outward flights
# vDeltaout = {}
# for t in range(0, pNperiods):
#     vDeltaout[t] = model.addVar(name='vDeltaout_%s' % (t), vtype=GRB.BINARY)

# # Binary auxiliary for applying or not the discount in return flights
# vDeltaret = {}
# for t in range(0, pNperiods):
#     vDeltaret[t] = model.addVar(name='vDeltaret_%s' % (t), vtype=GRB.BINARY)

# # Binary variable if person i works with role j
# vMu = {}
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         vMu[i,j] = model.addVar(ub=pProfiles[i,j], name='vMu_%s_%s' %(i,j), vtype=GRB.BINARY)

# # Number of people taking an outward flight with standard fare in time period t
# vXstand = {}
# for t in range(0, pNperiods):
#     vXstand[t] = model.addVar(lb=0.0, name='vXstand_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking an outward flight with discounted fare in time period t
# vXdisc = {}
# for t in range(0, pNperiods):
#     vXdisc[t] = model.addVar(lb=0.0, name='vXdisc_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a return flight with standard fare in time period t
# vYstand = {}
# for t in range(0, pNperiods):
#     vYstand[t] = model.addVar(lb=0.0, name='vYstand_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a return flight with discounted fare in time period t
# vYdisc = {}
# for t in range(0, pNperiods):
#     vYdisc[t] = model.addVar(lb=0.0, name='vYdisc_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a chartered flight outwards in period t
# vZout = {}
# for t in range(0, pNperiods):
#     vZout[t] = model.addVar(lb=0.0, name='vZout_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a chartered flight outwards in period t
# vZret = {}
# for t in range(0, pNperiods):
#     vZret[t] = model.addVar(lb=0.0, name='vZret_%s' % (t), vtype=GRB.INTEGER)

# # Slack variable
# vUmenos = {}
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods):
#         vUmenos[j,t] = model.addVar(lb=0.0, name='vUmenos_%s_%s' % (j,t), vtype=GRB.INTEGER)

# # Surplus variable
# vUmas = {}
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods):
#         vUmas[j,t] = model.addVar(lb=0.0, name='vUmas_%s_%s' % (j,t), vtype=GRB.INTEGER)

# model.setObjective(
#     pW1*(pWW1*quicksum(quicksum(pCostChar[t,l] * vGamma[t,l] for t in range(0,pNperiods)) for l in range(0,pNcharter)) +
#          pWW2*(quicksum(pCostOut[t] * vXstand[t] + (pCostOut[t] - pDiscount)* vXdisc[t] for t in range(0,pNperiods-1)) +
#          quicksum(pCostRet[t] * vYstand[t] + (pCostRet[t] - pDiscount) * vYdisc[t] for t in range(1,pNperiods)))) +
#     pW2*quicksum(quicksum(vUmas[j,t] for j in range(0,pNhealthp)) for t in range(0,pNperiods-1)) +
#     pW3*(pWW3*quicksum(quicksum(vMu[i,j] for i in range(0,pNpeople)) for j in range(0,pNhealthp)) +
#          pWW4*quicksum(quicksum(quicksum(pAvailability[i,t] * vBeta[i,j,t] for i in range(0, pNpeople)) for j in range(0, pNhealthp)) for t in range(0, pNperiods))),
#     sense=GRB.MINIMIZE)


# # CONSTRAINTS DEFINITION

# # Health profiles must be covered
# """
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods-1):
#         model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) >= pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))
# """
# # Health profiles must be covered treating infeasibility
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods-1):
#         model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) - vUmenos[j,t] + vUmas[j,t] == pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))

# # Minimum number of periods
# for i in range(0, pNpeople):
#     model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) >= pMinperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Minperiodsalpha_%s' %i)

# # Maximum number of periods
# for i in range(0, pNpeople):
#     model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) <= pMaxperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Maxperiodsalpha_%s' %i)

# # Maximum number of periods with beta
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) for t in range(0, pNperiods)) <= pMaxperiods, 'Maxperiodsbeta_%s' %i)

# # Controlling the outward and return flights for each person
# for i in range(0, pNpeople):
#     for t in range (0, pNperiods):
#         if t > 0:
#             model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) - quicksum(vBeta[i,j,t-1] for j in range(0, pNhealthp)) == vAlphaout[i,t] - vAlpharet[i,t], 'Flights_%s_%s' %(i,t))

# # Controlling the outward and return flights for each person in first period
# for i in range(0, pNpeople):
#     for t in range (0,1):
#         model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) == vAlphaout[i,t], 'Flights1_%s_%s' %(i,t))

# # One health profile for person at most for each time period
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         model.addConstr(quicksum(vBeta[i, j, t] for j in range(0, pNhealthp)) <= 1, 'Oneprofile_%s_%s' % (i, t))

# # Each person can fly once at most outward
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(vAlphaout[i, t] for t in range(0, pNperiods-1)) <= 1, 'Onceout_%s' % i)

# # Each person can fly once at most return
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(vAlpharet[i, t] for t in range(1, pNperiods)) <= 1, 'Onceret_%s' % i)

# # Number of outward are counted
# for t in range(0, pNperiods-1):
#     model.addConstr(vXstand[t] + vXdisc[t] + vZout[t] == quicksum(vAlphaout[i, t] for i in range(0, pNpeople)), 'Countout_%s' % t)

# # Number of return flights are counted
# for t in range(1, pNperiods):
#     model.addConstr(vYstand[t] + vYdisc[t] + vZret[t] == quicksum(vAlpharet[i, t] for i in range(0, pNpeople)), 'Countret_%s' % t)

# # Outward standard fare definition
# for t in range(0, pNperiods-1):
#     model.addConstr(vXstand[t] <= (pKpeople-1)*vDeltaout[t], 'StandOut_%s' % t)

# # Outward discounted fare definition 1
# for t in range(0, pNperiods-1):
#     model.addConstr(pKpeople*(1-vDeltaout[t]) <= vXdisc[t], 'DiscOut1_%s' % t)

# # Outward discounted fare definition 2
# for t in range(0, pNperiods-1):
#     model.addConstr(vXdisc[t] <= pNpeople*(1-vDeltaout[t]), 'DiscOut2_%s' % t)

# # Return standard fare definition
# for t in range(1, pNperiods):
#     model.addConstr(vYstand[t] <= (pKpeople-1)*vDeltaret[t], 'StandRet_%s' % t)

# # Return discounted fare definition 1
# for t in range(1, pNperiods):
#     model.addConstr(pKpeople*(1-vDeltaret[t]) <= vYdisc[t], 'DiscRet1_%s' % t)

# # Return discounted fare definition 2
# for t in range(1, pNperiods):
#     model.addConstr(vYdisc[t] <= pNpeople*(1-vDeltaret[t]), 'DiscRet2_%s' % t)

# # Only one type of chartered flight is hired at most
# for t in range(0, pNperiods):
#     model.addConstr(quicksum(vGamma[t,l] for l in range(0,pNcharter)) <= 1, 'Onechtype_%s' % t)

# # Satisfying minimum capacity of chartered in outward flights
# for t in range(0, pNperiods - 1):
#     model.addConstr(quicksum(pMinCapCh[l] * vGamma[t,l] for l in range(0,pNcharter)) <= vZout[t], 'Mincapout_%s' % t)

# # Satisfying maximum capacity of chartered in outward flights
# for t in range(0, pNperiods - 1):
#     model.addConstr(vZout[t] <= quicksum(pMaxCapCh[l] * vGamma[t,l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# # Satisfying minimum capacity of chartered in return flights
# for t in range(1, pNperiods):
#     model.addConstr(quicksum(pMinCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)) <= vZret[t], 'Mincapout_%s' % t)

# # Satisfying maximum capacity of chartered in return flights
# for t in range(1, pNperiods):
#     model.addConstr(vZret[t] <= quicksum(pMaxCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# # Number of roles that a person plays 1
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) <= pMaxperiods * vMu[i,j], 'Roles1_%s_%s' %(i,j))

# # Number of roles that a person plays 2
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) >= vMu[i,j], 'Roles1_%s_%s' %(i,j))

class STARTmodel:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.data = pd.read_excel(excel_file, sheet_name='Data', header=None)
        self.pNperiods = self.data.loc[0,1]
        self.pNhealthp = self.data.loc[1,1]
        self.pNpeople = self.data.loc[2,1]
        self.pNcharter = self.data.loc[3,1]
        self.pDiscount = self.data.loc[4,1]
        self.pKpeople = self.data.loc[5,1]
        self.pMaxperiods = self.data.loc[6,1]
        self.pMinperiods = self.data.loc[7,1]

        self.prices = pd.read_excel(excel_file, sheet_name='Prices')
        self.pCostOut = self.prices.loc[:,'Outward'].to_numpy()
        self.pCostRet = self.prices.loc[:,'Return'].to_numpy()

        self.charter = pd.read_excel(excel_file, sheet_name='Charter')
        self.pCostChar = self.charter.loc[0:self.nperiods-1,'Chartered 1':'Chartered 2'].to_numpy()
        self.pMinCapCh = self.charter.loc[0:self.ncharter-1,'Min Cap'].to_numpy()
        self.pMaxCapCh = self.charter.loc[0:self.ncharter-1,'Max Cap'].to_numpy()

        self.demand = pd.read_excel(excel_file, sheet_name='Demand')
        self.pHealthdem = self.demand.loc[:,1:].to_numpy()

        self.healthprofiles = pd.read_excel(excel_file, sheet_name='HealthProfiles', header=None)
        self.pNameprofiles = self.healthprofiles.loc[0,1:].to_numpy()
        self.pNameabbprofiles = self.healthprofiles.loc[1,1:].to_numpy()
        self.pProfiles = self.healthprofiles.loc[2:,1:].to_numpy()

        self.availability = pd.read_excel(excel_file, sheet_name='Availability')
        self.pAvailability = self.availability.loc[:,1:].to_numpy()

        self.weights = pd.read_excel(excel_file, sheet_name='Weights', header=None)
        self.pW1 = self.weights.loc[0,1]
        self.pW2 = self.weights.loc[1,1]
        self.pW3 = self.weights.loc[2,1]

        self.pWW1 = self.weights.loc[4,1]
        self.pWW2 = self.weights.loc[5,1]
        self.pWW3 = self.weights.loc[6,1]
        self.pWW4 = self.weights.loc[7,1]

        self.vApplhaout = {}
        self.vAlpharet = {}
        self.vBeta = {}
        self.vGamma = {}
        self.vDeltaout = {}
        self.vDeltaret = {}
        self.vMu = {}
        self.vXstand = {}
        self.vXdisc = {}
        self.vYstand = {}
        self.vYdisc = {}
        self.vZout = {}
        self.vZret = {}
        self.vUmenos = {}
        self.vUmas = {}
        
    def create_model(self):
        self.model = Model("START")
        return self.model

    def create_variables(self): 
        for i in range(0, self.pNpeople):
             for t in range(0, self.pNperiods):
                self.vAlphaout[i,t] = self.model.addVar(name='vAlphaout_%s_%s' % (i, t), vtype=GRB.BINARY)
        
        for i in range(0, self.pNpeople):
            for t in range(0, self.pNperiods):
                self.vAlpharet[i,t] = self.model.addVar(name='vAlpharet_%s_%s' % (i, t), vtype=GRB.BINARY)
        
        for i in range(0, self.pNpeople):
            for j in range(0, self.pNhealthp):
                for t in range(0, self.pNperiods):
                    self.vBeta[i, j, t] = self.model.addVar(ub=min(self.pAvailability[i,t],self.pProfiles[i,j]), name='vBeta_%s_%s_%s' % (i, j, t), vtype=GRB.BINARY)

        for t in range(0, self.pNperiods):
            for l in range(0, self.pNcharter):
                self.vGamma[t,l] = self.model.addVar(name='vGamma_%s_%s' % (t,l), vtype=GRB.BINARY)
        
        for t in range(0, self.pNperiods):
            self.vDeltaout[t] = self.model.addVar(name='vDeltaout_%s' % (t), vtype=GRB.BINARY)

        for t in range(0, self.pNperiods):
            self.vDeltaret[t] = self.model.addVar(name='vDeltaret_%s' % (t), vtype=GRB.BINARY)
        
        for i in range(0, self.pNpeople):
            for j in range(0, self.pNhealthp):
                self.vMu[i,j] = self.model.addVar(ub=self.pProfiles[i,j], name='vMu_%s_%s' %(i,j), vtype=GRB.BINARY)

        for t in range(0, self.pNperiods):
            self.vXstand[t] = self.model.addVar(lb=0.0, name='vXstand_%s' % (t), vtype=GRB.INTEGER)

        for t in range(0, self.pNperiods):
            self.vXdisc[t] = self.model.addVar(lb=0.0, name='vXdisc_%s' % (t), vtype=GRB.INTEGER)
        
        for t in range(0, self.pNperiods):
            self.vYstand[t] = self.model.addVar(lb=0.0, name='vYstand_%s' % (t), vtype=GRB.INTEGER)
        
        for t in range(0, self.pNperiods):
            self.vYdisc[t] = self.model.addVar(lb=0.0, name='vYdisc_%s' % (t), vtype=GRB.INTEGER)
        
        for t in range(0, self.pNperiods):
            self.vZout[t] = self.model.addVar(lb=0.0, name='vZout_%s' % (t), vtype=GRB.INTEGER)
        
        for t in range(0, self.pNperiods):
            self.vZret[t] = self.model.addVar(lb=0.0, name='vZret_%s' % (t), vtype=GRB.INTEGER)
        
        for j in range(0, self.pNhealthp):
            for t in range(0, self.pNperiods):
                self.vUmenos[j,t] = self.model.addVar(lb=0.0, name='vUmenos_%s_%s' % (j,t), vtype=GRB.INTEGER)
        
        for j in range(0, self.pNhealthp):
            for t in range(0, self.pNperiods):
                self.vUmas[j,t] = self.model.addVar(lb=0.0, name='vUmas_%s_%s' % (j,t), vtype=GRB.INTEGER)
        

    def create_objective_function(self):
        self.model.setObjective(
            self.pW1*(self.pWW1*quicksum(quicksum(self.pCostChar[t,l] * self.vGamma[t,l] for t in range(0,self.pNperiods)) for l in range(0,self.pNcharter)) +
            self.pWW2*(quicksum(self.pCostOut[t] * self.vXstand[t] + (self.pCostOut[t] - self.pDiscount)* self.vXdisc[t] for t in range(0,self.pNperiods-1)) +
            quicksum(self.pCostRet[t] * self.vYstand[t] + (self.pCostRet[t] - self.pDiscount) * self.vYdisc[t] for t in range(1,self.pNperiods)))) +
            self.pW2*quicksum(quicksum(self.vUmas[j,t] for j in range(0,self.pNhealthp)) for t in range(0,self.pNperiods-1)) +
            self.pW3*(self.pWW3*quicksum(quicksum(self.vMu[i,j] for i in range(0,self.pNpeople)) for j in range(0,self.pNhealthp)) +
            self.pWW4*quicksum(quicksum(quicksum(self.pAvailability[i,t] * self.vBeta[i,j,t] for i in range(0, self.pNpeople)) for j in range(0, self.pNhealthp)) for t in range(0, self.pNperiods))),
            sense=GRB.MINIMIZE)
    
    def create_constrains(self):
        for j in range(0, self.pNhealthp):
            for t in range(0, self.pNperiods-1):
                self.model.addConstr(quicksum(self.pProfiles[i,j]*self.vBeta[i,j,t] for i in range(0, self.pNpeople)) - self.vUmenos[j,t] + self.vUmas[j,t] == self.pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))

        for i in range(0, self.pNpeople):
            self.model.addConstr(quicksum((t+1)*self.vAlpharet[i,t] for t in range(1, self.pNperiods)) - quicksum((t+1)*self.vAlphaout[i,t] for t in range(0,self.pNperiods-1)) >= self.pMinperiods * quicksum(self.vAlphaout[i, t] for t in range(0, self.pNperiods)), 'Minperiodsalpha_%s' %i)
        
        for i in range(0, self.pNpeople):
            self.model.addConstr(quicksum((t+1)*self.vAlpharet[i,t] for t in range(1, self.pNperiods)) - quicksum((t+1)*self.vAlphaout[i,t] for t in range(0,self.pNperiods-1)) <= self.pMaxperiods * quicksum(self.vAlphaout[i, t] for t in range(0, self.pNperiods)), 'Maxperiodsalpha_%s' %i)

        for i in range(0, self.pNpeople):
            self.model.addConstr(quicksum(quicksum(self.vBeta[i,j,t] for j in range(0, self.pNhealthp)) for t in range(0, self.pNperiods)) <= self.pMaxperiods, 'Maxperiodsbeta_%s' %i)
        
        for i in range(0, self.pNpeople):
            for t in range (0, self.pNperiods):
                if t > 0:
                    self.model.addConstr(quicksum(self.vBeta[i,j,t] for j in range(0, self.pNhealthp)) - quicksum(self.vBeta[i,j,t-1] for j in range(0, self.pNhealthp)) == self.vAlphaout[i,t] - self.vAlpharet[i,t], 'Flights_%s_%s' %(i,t))
        
        for i in range(0, self.pNpeople):
            for t in range (0,1):
                self.model.addConstr(quicksum(self.vBeta[i,j,t] for j in range(0, self.pNhealthp)) == self.vAlphaout[i,t], 'Flights1_%s_%s' %(i,t))
        
        for i in range(0, self.pNpeople):
            for t in range(0, self.pNperiods):
                self.model.addConstr(quicksum(self.vBeta[i, j, t] for j in range(0, self.pNhealthp)) <= 1, 'Oneprofile_%s_%s' % (i, t))
        
        for i in range(0, self.pNpeople):
            self.model.addConstr(quicksum(self.vAlphaout[i, t] for t in range(0, self.pNperiods-1)) <= 1, 'Onceout_%s' % i)

        for i in range(0, self.pNpeople):
            self.model.addConstr(quicksum(self.vAlpharet[i, t] for t in range(1, self.pNperiods)) <= 1, 'Onceret_%s' % i)
        
        for t in range(0, self.pNperiods-1):
            self.model.addConstr(self.vXstand[t] + self.vXdisc[t] + self.vZout[t] == quicksum(self.vAlphaout[i, t] for i in range(0, self.pNpeople)), 'Countout_%s' % t)

        for t in range(1, self.pNperiods):
            self.model.addConstr(self.vYstand[t] + self.vYdisc[t] + self.vZret[t] == quicksum(self.vAlpharet[i, t] for i in range(0, self.pNpeople)), 'Countret_%s' % t)
        
        for t in range(0, self.pNperiods-1):
            self.model.addConstr(self.vXstand[t] <= (self.pKpeople-1)*self.vDeltaout[t], 'StandOut_%s' % t)
        
        for t in range(0, self.pNperiods-1):
            self.model.addConstr(self.pKpeople*(1-self.vDeltaout[t]) <= self.vXdisc[t], 'DiscOut1_%s' % t)

        for t in range(0, self.pNperiods-1):
            self.model.addConstr(self.vXdisc[t] <= self.pNpeople*(1-self.vDeltaout[t]), 'DiscOut2_%s' % t)
        
        for t in range(1, self.pNperiods):
            self.model.addConstr(self.vYstand[t] <= (self.pKpeople-1)*self.vDeltaret[t], 'StandRet_%s' % t)
        
        for t in range(1, self.pNperiods):
            self.model.addConstr(self.pKpeople*(1-self.vDeltaret[t]) <= self.vYdisc[t], 'DiscRet1_%s' % t)
        
        for t in range(1, self.pNperiods):
            self.model.addConstr(self.vYdisc[t] <= self.pNpeople*(1-self.vDeltaret[t]), 'DiscRet2_%s' % t)
        
        for t in range(0, self.pNperiods):
            self.model.addConstr(quicksum(self.vGamma[t,l] for l in range(0,self.pNcharter)) <= 1, 'Onechtype_%s' % t)
        
        for t in range(0, self.pNperiods - 1):
            self.model.addConstr(quicksum(self.pMinCapCh[l] * self.vGamma[t,l] for l in range(0,self.pNcharter)) <= self.vZout[t], 'Mincapout_%s' % t)
        
        for t in range(0, self.pNperiods - 1):
            self.model.addConstr(self.vZout[t] <= quicksum(self.pMaxCapCh[l] * self.vGamma[t,l] for l in range(0, self.pNcharter)), 'Mincapout_%s' % t)
        
        for t in range(1, self.pNperiods):
            self.model.addConstr(quicksum(self.pMinCapCh[l] * self.vGamma[t, l] for l in range(0, self.pNcharter)) <= self.vZret[t], 'Mincapout_%s' % t)
        
        for t in range(1, self.pNperiods):
            self.model.addConstr(self.vZret[t] <= quicksum(self.pMaxCapCh[l] * self.vGamma[t, l] for l in range(0, self.pNcharter)), 'Mincapout_%s' % t)
        
        for i in range(0, self.pNpeople):
            for j in range(0, self.pNhealthp):
                self.model.addConstr(quicksum(self.vBeta[i,j,t] for t in range(0,self.pNperiods-1)) <= self.pMaxperiods * self.vMu[i,j], 'Roles1_%s_%s' %(i,j))
        
        for i in range(0, self.pNpeople):
            for j in range(0, self.pNhealthp):
                self.model.addConstr(quicksum(self.vBeta[i,j,t] for t in range(0,self.pNperiods-1)) >= self.vMu[i,j], 'Roles1_%s_%s' %(i,j))
            
    def solve(self):
        self.model.optimize()
        return self.model



# # Abstract object for the mathematical optimisation model
# model = Model("START")

# # VARIABLES DEFINITION
# # Binary for person i taking an outward flight in time period t
# vAlphaout = {}
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         vAlphaout[i,t] = model.addVar(name='vAlphaout_%s_%s' % (i, t), vtype=GRB.BINARY)

# # Binary for person i taking a return flight in time period t
# vAlpharet = {}
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         vAlpharet[i,t] = model.addVar(name='vAlpharet_%s_%s' % (i, t), vtype=GRB.BINARY)

# # Binary for person i attending with profile j in time period t
# vBeta = {}
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         for t in range(0, pNperiods):
#             vBeta[i, j, t] = model.addVar(ub=min(pAvailability[i,t],pProfiles[i,j]), name='vBeta_%s_%s_%s' % (i, j, t), vtype=GRB.BINARY)

# # Binary for chartered flight type 1 in period t
# vGamma = {}
# for t in range(0, pNperiods):
#     for l in range(0, pNcharter):
#         vGamma[t,l] = model.addVar(name='vGamma_%s_%s' % (t,l), vtype=GRB.BINARY)

# # Binary auxiliary for applying or not the discount in outward flights
# vDeltaout = {}
# for t in range(0, pNperiods):
#     vDeltaout[t] = model.addVar(name='vDeltaout_%s' % (t), vtype=GRB.BINARY)

# # Binary auxiliary for applying or not the discount in return flights
# vDeltaret = {}
# for t in range(0, pNperiods):
#     vDeltaret[t] = model.addVar(name='vDeltaret_%s' % (t), vtype=GRB.BINARY)

# # Binary variable if person i works with role j
# vMu = {}
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         vMu[i,j] = model.addVar(ub=pProfiles[i,j], name='vMu_%s_%s' %(i,j), vtype=GRB.BINARY)

# # Number of people taking an outward flight with standard fare in time period t
# vXstand = {}
# for t in range(0, pNperiods):
#     vXstand[t] = model.addVar(lb=0.0, name='vXstand_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking an outward flight with discounted fare in time period t
# vXdisc = {}
# for t in range(0, pNperiods):
#     vXdisc[t] = model.addVar(lb=0.0, name='vXdisc_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a return flight with standard fare in time period t
# vYstand = {}
# for t in range(0, pNperiods):
#     vYstand[t] = model.addVar(lb=0.0, name='vYstand_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a return flight with discounted fare in time period t
# vYdisc = {}
# for t in range(0, pNperiods):
#     vYdisc[t] = model.addVar(lb=0.0, name='vYdisc_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a chartered flight outwards in period t
# vZout = {}
# for t in range(0, pNperiods):
#     vZout[t] = model.addVar(lb=0.0, name='vZout_%s' % (t), vtype=GRB.INTEGER)

# # Number of people taking a chartered flight outwards in period t
# vZret = {}
# for t in range(0, pNperiods):
#     vZret[t] = model.addVar(lb=0.0, name='vZret_%s' % (t), vtype=GRB.INTEGER)

# # Slack variable
# vUmenos = {}
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods):
#         vUmenos[j,t] = model.addVar(lb=0.0, name='vUmenos_%s_%s' % (j,t), vtype=GRB.INTEGER)

# # Surplus variable
# vUmas = {}
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods):
#         vUmas[j,t] = model.addVar(lb=0.0, name='vUmas_%s_%s' % (j,t), vtype=GRB.INTEGER)


# # OBJECTIVE FUNCTION
# # Objective function definition treating infeasibility
# model.setObjective(
#     pW1*(pWW1*quicksum(quicksum(pCostChar[t,l] * vGamma[t,l] for t in range(0,pNperiods)) for l in range(0,pNcharter)) +
#          pWW2*(quicksum(pCostOut[t] * vXstand[t] + (pCostOut[t] - pDiscount)* vXdisc[t] for t in range(0,pNperiods-1)) +
#          quicksum(pCostRet[t] * vYstand[t] + (pCostRet[t] - pDiscount) * vYdisc[t] for t in range(1,pNperiods)))) +
#     pW2*quicksum(quicksum(vUmas[j,t] for j in range(0,pNhealthp)) for t in range(0,pNperiods-1)) +
#     pW3*(pWW3*quicksum(quicksum(vMu[i,j] for i in range(0,pNpeople)) for j in range(0,pNhealthp)) +
#          pWW4*quicksum(quicksum(quicksum(pAvailability[i,t] * vBeta[i,j,t] for i in range(0, pNpeople)) for j in range(0, pNhealthp)) for t in range(0, pNperiods))),
#     sense=GRB.MINIMIZE)

# # CONSTRAINTS DEFINITION

# # Health profiles must be covered
# """
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods-1):
#         model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) >= pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))
# """
# # Health profiles must be covered treating infeasibility
# for j in range(0, pNhealthp):
#     for t in range(0, pNperiods-1):
#         model.addConstr(quicksum(pProfiles[i,j]*vBeta[i,j,t] for i in range(0, pNpeople)) - vUmenos[j,t] + vUmas[j,t] == pHealthdem[j,t], 'Profiles_%s_%s' %(j,t))

# # Minimum number of periods
# for i in range(0, pNpeople):
#     model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) >= pMinperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Minperiodsalpha_%s' %i)

# # Maximum number of periods
# for i in range(0, pNpeople):
#     model.addConstr(quicksum((t+1)*vAlpharet[i,t] for t in range(1, pNperiods)) - quicksum((t+1)*vAlphaout[i,t] for t in range(0,pNperiods-1)) <= pMaxperiods * quicksum(vAlphaout[i, t] for t in range(0, pNperiods)), 'Maxperiodsalpha_%s' %i)

# # Maximum number of periods with beta
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) for t in range(0, pNperiods)) <= pMaxperiods, 'Maxperiodsbeta_%s' %i)

# # Controlling the outward and return flights for each person
# for i in range(0, pNpeople):
#     for t in range (0, pNperiods):
#         if t > 0:
#             model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) - quicksum(vBeta[i,j,t-1] for j in range(0, pNhealthp)) == vAlphaout[i,t] - vAlpharet[i,t], 'Flights_%s_%s' %(i,t))

# # Controlling the outward and return flights for each person in first period
# for i in range(0, pNpeople):
#     for t in range (0,1):
#         model.addConstr(quicksum(vBeta[i,j,t] for j in range(0, pNhealthp)) == vAlphaout[i,t], 'Flights1_%s_%s' %(i,t))

# # One health profile for person at most for each time period
# for i in range(0, pNpeople):
#     for t in range(0, pNperiods):
#         model.addConstr(quicksum(vBeta[i, j, t] for j in range(0, pNhealthp)) <= 1, 'Oneprofile_%s_%s' % (i, t))

# # Each person can fly once at most outward
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(vAlphaout[i, t] for t in range(0, pNperiods-1)) <= 1, 'Onceout_%s' % i)

# # Each person can fly once at most return
# for i in range(0, pNpeople):
#     model.addConstr(quicksum(vAlpharet[i, t] for t in range(1, pNperiods)) <= 1, 'Onceret_%s' % i)

# # Number of outward are counted
# for t in range(0, pNperiods-1):
#     model.addConstr(vXstand[t] + vXdisc[t] + vZout[t] == quicksum(vAlphaout[i, t] for i in range(0, pNpeople)), 'Countout_%s' % t)

# # Number of return flights are counted
# for t in range(1, pNperiods):
#     model.addConstr(vYstand[t] + vYdisc[t] + vZret[t] == quicksum(vAlpharet[i, t] for i in range(0, pNpeople)), 'Countret_%s' % t)

# # Outward standard fare definition
# for t in range(0, pNperiods-1):
#     model.addConstr(vXstand[t] <= (pKpeople-1)*vDeltaout[t], 'StandOut_%s' % t)

# # Outward discounted fare definition 1
# for t in range(0, pNperiods-1):
#     model.addConstr(pKpeople*(1-vDeltaout[t]) <= vXdisc[t], 'DiscOut1_%s' % t)

# # Outward discounted fare definition 2
# for t in range(0, pNperiods-1):
#     model.addConstr(vXdisc[t] <= pNpeople*(1-vDeltaout[t]), 'DiscOut2_%s' % t)

# # Return standard fare definition
# for t in range(1, pNperiods):
#     model.addConstr(vYstand[t] <= (pKpeople-1)*vDeltaret[t], 'StandRet_%s' % t)

# # Return discounted fare definition 1
# for t in range(1, pNperiods):
#     model.addConstr(pKpeople*(1-vDeltaret[t]) <= vYdisc[t], 'DiscRet1_%s' % t)

# # Return discounted fare definition 2
# for t in range(1, pNperiods):
#     model.addConstr(vYdisc[t] <= pNpeople*(1-vDeltaret[t]), 'DiscRet2_%s' % t)

# # Only one type of chartered flight is hired at most
# for t in range(0, pNperiods):
#     model.addConstr(quicksum(vGamma[t,l] for l in range(0,pNcharter)) <= 1, 'Onechtype_%s' % t)

# # Satisfying minimum capacity of chartered in outward flights
# for t in range(0, pNperiods - 1):
#     model.addConstr(quicksum(pMinCapCh[l] * vGamma[t,l] for l in range(0,pNcharter)) <= vZout[t], 'Mincapout_%s' % t)

# # Satisfying maximum capacity of chartered in outward flights
# for t in range(0, pNperiods - 1):
#     model.addConstr(vZout[t] <= quicksum(pMaxCapCh[l] * vGamma[t,l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# # Satisfying minimum capacity of chartered in return flights
# for t in range(1, pNperiods):
#     model.addConstr(quicksum(pMinCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)) <= vZret[t], 'Mincapout_%s' % t)

# # Satisfying maximum capacity of chartered in return flights
# for t in range(1, pNperiods):
#     model.addConstr(vZret[t] <= quicksum(pMaxCapCh[l] * vGamma[t, l] for l in range(0, pNcharter)), 'Mincapout_%s' % t)

# # Number of roles that a person plays 1
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) <= pMaxperiods * vMu[i,j], 'Roles1_%s_%s' %(i,j))

# # Number of roles that a person plays 2
# for i in range(0, pNpeople):
#     for j in range(0, pNhealthp):
#         model.addConstr(quicksum(vBeta[i,j,t] for t in range(0,pNperiods-1)) >= vMu[i,j], 'Roles1_%s_%s' %(i,j))

# # Writing model in format mps
# model.write('model.mps')

