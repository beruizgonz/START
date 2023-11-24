from gurobipy import *

# Create a model
m = Model()

# Create variables
x = m.addVar(name="x")
y = m.addVar(name="y")

# Set objective function: Maximize 3x + 2y
m.setObjective(3*x + 2*y, GRB.MAXIMIZE)

# Add constraints
m.addConstr(2*x + y <= 20, "c1")
m.addConstr(4*x - 5*y >= -10, "c2")
m.addConstr(x >= 0, "c3")
m.addConstr(y >= 3, "c4")

# Optimize the model
m.optimize()

# Print the optimal values of the decision variables
print(f"Optimal value of x: {x.x}")
print(f"Optimal value of y: {y.x}")

# Print the optimal objective value
print(f"Optimal objective value: {m.objVal}")
