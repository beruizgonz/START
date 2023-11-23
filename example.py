import gurobipy as gp
from gurobipy import GRB

# Create a new model 
model = gp.Model("flight_assignment")
a = 1
# Sets
T = ['t1', 't2', 't3', 't4', 't5', 't6']
J = ['j1', 'j2', 'j3', 'j4', 'j5', 'j6']
I = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']

# Parameters
ce = {'t1': 1581, 't2': 953, 't3': 696, 't4': 592, 't5': 592, 't6': 0}
cv = {'t1': 845, 't2': 973, 't3': 973, 't4': 1072, 't5': 615, 't6': 885}
n = {'j1': 2, 'j2': 1, 'j3': 1, 'j4': 1, 'j5': 2, 'j6': 2}

p = {
    '1': {'j1': 1, 'j2': 0, 'j3': 0, 'j4': 0, 'j5': 0, 'j6': 0},
    '2': {'j1': 0, 'j2': 0, 'j3': 1, 'j4': 0, 'j5': 0, 'j6': 1},
    '3': {'j1': 0, 'j2': 0, 'j3': 1, 'j4': 0, 'j5': 0, 'j6': 1},
    # ... (complete the dictionary with the remaining entries)
}

d = {
    '1': {'t1': 1, 't2': 1, 't3': 1, 't4': 0, 't5': 0, 't6': 0},
    '2': {'t1': 0, 't2': 1, 't3': 1, 't4': 0, 't5': 0, 't6': 0},
    '3': {'t1': 1, 't2': 0, 't3': 1, 't4': 1, 't5': 1, 't6': 0},
    # ... (complete the dictionary with the remaining entries)
}

# Decision Variables
# Assuming 'i', 'j', and 't' are iterable objects or ranges
x = {(i, j, t): model.addVar(vtype=GRB.BINARY,name='x_{}_{}_{}'.format(i, j, t))
     for i in i  # Replace "..." with the appropriate range or iterable for i
     for j in j # Replace "..." with the appropriate range or iterable for j
     for t in t}  # Replace "..." with the appropriate range or iterable for t


# Objective Function
model.setObjective(
    gp.quicksum(ce[t] * x[i, j, t] for i in i for j in j for t in t) +
    gp.quicksum(cv[t] * x[i, j, t] for i in i for j in j for t in t),
    GRB.MINIMIZE
)

# Constraints
for j_val in j:
    model.addConstr(gp.quicksum(n[j_val] * x[i, j_val, t] for i in i for t in t) >= n[j_val], f'constraint_{j_val}')

for i_val in i:
    for t_val in t:
        model.addConstr(gp.quicksum(p[i_val][j_val] * x[i_val, j_val, t_val] for j_val in j) <= d[i_val][t_val],
                        f'constraint_{i_val}_{t_val}')

# Optimize the model
model.optimize()

# Print the optimal solution
if model.status == GRB.OPTIMAL:
    for i_val in i:
        for j_val in j:
            for t_val in t:
                if x[i_val, j_val, t_val].x > 0.5:
                    print(f'Assign {n[j_val]} people of profile {j_val} to period {t_val} for person {i_val}')

# Print the objective value
print(f'Optimal objective value: {model.objVal}')