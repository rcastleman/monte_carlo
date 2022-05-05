# import xlwings as xw
# from xlwings import constants
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
import matplotlib.pyplot as plt
from datetime import date,timedelta
from pyxirr import xirr


# Returns Distribution Parameters (lognormal distribution)
returns_mean = 1.0
returns_STD = 1.0
# returns_sims = 100000

# plt.hist(np.random.lognormal(returns_mean,returns_STD,returns_sims),
#         bins = 500,
#         density = True,
#         align = 'mid')
# plt.xlim(0,50)
# plt.show()

# Investment Period Distribution Parameters (triangular distribution)

inv_period_lower_bound = 0
inv_period_mode = 2.0
inv_period_upper_bound = 5
inv_period_sims = 100000
inv_period_bins = inv_period_upper_bound

# plt.hist(np.random.triangular(inv_period_lower_bound,inv_period_mode,inv_period_upper_bound,inv_period_sims),
#         bins = inv_period_bins,
#         density = True,
#         align = 'mid')
# plt.xlim(0,6)
# plt.show()

# Hold Period Distribution Parameters (triangular distribution)

hold_period_sims = 1000
hold_period_mean= 4.5

hold_period_lower_bound = 1
hold_period_mode = hold_period_mean
hold_period_upper_bound = 10

# hold_period_normal_dist = np.random.triangular(hold_period_lower_bound,
#                                                hold_period_mode,
#                                                hold_period_upper_bound,
#                                                hold_period_sims)

# plt.hist(hold_period_normal_dist,
#         bins = 100,
#         density = True,
#         align = 'mid')
# plt.xlim(0,hold_period_upper_bound)
# plt.show()

# CREATE A PORTFOLIO

class Company:
    def __init__(self):
        self.capital_in = None
        self.capital_out = None 
        self.inv_date = None
        self.exit_date = None
        self.hold_period = None
        self.MOIC = None

# fund_size = 100
# num_companies = 5
# avg_initial_investment = fund_size / num_companies
inv_period_begin_date = date(2022,1,1)
days_in_year = 365
results = []
hold_period_mode = 4.5
hold_period_lower_bound = 1
hold_period_upper_bound = 10

dates = []
amounts = []

def portfolio():

    avg_initial_investment = fund_size / num_companies

    # Create all company objects
    companies = [Company() for _ in range(num_companies)]
    dates.clear()
    amounts.clear()

    
    for company in companies:
        company.capital_in = avg_initial_investment  # all initial investments are currently the Fund Size divided by the number of companies 
        
        company.capital_out = company.capital_in * np.random.lognormal(returns_mean,returns_STD) #see parameters above
        
        company.inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound * days_in_year,
        inv_period_mode * days_in_year,
        inv_period_upper_bound * days_in_year)) #random offset from Investment Period Start Date using parameters above 
        
        company.exit_date = company.inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound * days_in_year,
        hold_period_mode * days_in_year, 
        hold_period_upper_bound * days_in_year)) #random offset from Company Investment Date using parameters above

        dates.append(company.inv_date)
        dates.append(company.exit_date)
        amounts.append(-company.capital_in)
        amounts.append(company.capital_out)

        company.hold_period = company.exit_date - company.inv_date
        # company.MOIC = company.capital_out / company.capital_in

    portfolio_IRR = xirr(dates,amounts)
    return portfolio_IRR
        
# portfolio()
# print(xirr(dates,amounts))

#SIMULATION N PORTFOLIOS AND RETURN RANGE OF OUTCOMES 

results = []
fund_size = int(input ("Enter fund size: "))
num_companies = int(input("Enter # of companies in fund: "))
num_sims = int(input("Enter # of simulations: "))

for i in range(num_sims):
    results.append(portfolio())

# print(results)

output_data = pd.DataFrame(results,columns = ["Simulation #","Fund Size","# Companies","Holding Period","MOIC","IRR"])
print(output_data)