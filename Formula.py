# rendering the XIRR formula in code

# the goal is to iterate through a succession of IRRs such that the formula approaches zero:

# sum of ... Payment / (1+XIRR)^((date[i]  - date [0])/365)

# see https://github.com/rcastleman/monte_carlo/blob/main/formula.png

# tolerance = range between which, if answers vary, formula is "solved"

xirr = 0.1 
dates_in = []
dates_out = []
capital_in = []
capital_out = []

