import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
from gurobipy import quicksum, Model, GRB

class START:
    def __init__(self, excel_file):
        super(START).__init__()
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
        self.pCostChar = self.charter.loc[0:self.pNperiods-1,'Chartered 1':'Chartered 2'].to_numpy()
        self.pMinCapCh = self.charter.loc[0:self.pNcharter-1,'Min Cap'].to_numpy()
        self.pMaxCapCh = self.charter.loc[0:self.pNcharter-1,'Max Cap'].to_numpy()

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

        self.vAlphaout = {}
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





