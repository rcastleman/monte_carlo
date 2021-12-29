import xlwings as xw
from xlwings import constants
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
import matplotlib.pyplot as plt
from datetime import date,timedelta

# returns distribution parameters
returns_mean = 1.0
returns_STD = 1.0
returns_sims = 100_000

plt.hist(np.random.lognormal(returns_mean,returns_STD,returns_sims),
        bins = 500,
        density = True,
        align = 'mid')
plt.xlim(0,50)
plt.show()


# investment period distribution parameters

inv_period_lower_bound = 0
inv_period_mode = 2.0
inv_period_upper_bound = 5
inv_period_sims = 100_000
inv_period_bins = inv_period_upper_bound

plt.hist(np.random.triangular(inv_period_lower_bound,inv_period_mode,inv_period_upper_bound,inv_period_sims),
        bins = inv_period_bins,
        density = True,
        align = 'mid')
plt.xlim(0,6)
plt.show()

# hold period distribution parameters

hold_period_sims = 1000
hold_period_mean= 4.5

#Normal Distribution parameters

# hold_period_STD = 2.0
# plt.hist(np.random.normal(hold_period_mean,hold_period_STD,hold_period_sims),
#         bins = 1_000,
#         density = True,
#         align = 'mid')
# plt.xlim(0,upper_bound)
# plt.show()

#Triangular Distribution
hold_period_lower_bound = 1
hold_period_mode = hold_period_mean
hold_period_upper_bound = 10

hold_period_normal_dist = np.random.triangular(hold_period_lower_bound,
                                               hold_period_mode,
                                               hold_period_upper_bound,
                                               hold_period_sims)

plt.hist(hold_period_normal_dist,
        bins = 100,
        density = True,
        align = 'mid')
plt.xlim(0,hold_period_upper_bound)
plt.show()


# Connect the Excel workbook

book = xw.Book('portfolio_monte_carlo.xlsx')
model = book.sheets("Model")
results = book.sheets("Results")


# SIMULATION

fund_size = 100
num_companies = 10
inv_period_begin_date = date(2022,1,1)
days = 365
results = []
hold_period_mode = 4.5
hold_period_lower_bound = 1
hold_period_upper_bound = 10

def dcf_simulation(fund_size,
                   num_companies,
                   inv_period_begin_date,
                   days,
                   results,
                   inv_period_lower_bound,
                   inv_period_mode,
                   inv_period_upper_bound,
                   hold_period_lower_bound,
                   hold_period_mode,
                   hold_period_upper_bound):
    
    book.sheets('Results').clear()

#     model.range("C2").value = fund_size
#     model.range("C3").value = num_companies

    avg_initial_investment = fund_size / num_companies

    #capital in
    model.range("C16").value = -avg_initial_investment
    model.range("C17").value = -avg_initial_investment
    model.range("C18").value = -avg_initial_investment
    model.range("C19").value = -avg_initial_investment
    model.range("C20").value = -avg_initial_investment
    model.range("C21").value = -avg_initial_investment
    model.range("C22").value = -avg_initial_investment
    model.range("C23").value = -avg_initial_investment
    model.range("C24").value = -avg_initial_investment
    model.range("C25").value = -avg_initial_investment

    #capital out
    company1_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D16").value = company1_capital_out
    company2_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D17").value = company2_capital_out
    company3_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D18").value = company3_capital_out
    company4_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D19").value = company4_capital_out
    company5_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D20").value = company5_capital_out
    company6_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D21").value = company6_capital_out
    company7_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D22").value = company7_capital_out
    company8_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D23").value = company8_capital_out
    company9_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D24").value = company9_capital_out
    company10_capital_out = avg_initial_investment * np.random.lognormal(returns_mean,returns_STD)
    model.range("D25").value = company10_capital_out
    
    #capital in date
    company1_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E16").value = company1_inv_date
    
    company2_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E17").value = company2_inv_date
    
    company3_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E18").value = company3_inv_date
    
    company4_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E19").value = company4_inv_date
    
    company5_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E20").value = company5_inv_date
    
    company6_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E21").value = company6_inv_date
    
    company7_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E22").value = company7_inv_date
    
    company8_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E23").value = company8_inv_date
    
    company9_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E24").value = company9_inv_date
    
    company10_inv_date = inv_period_begin_date + timedelta(days = np.random.triangular(inv_period_lower_bound*days,
                                                         inv_period_mode*days,
                                                         inv_period_upper_bound*days))
    model.range("E25").value = company10_inv_date
    
    
    

    #capital out date
    company1_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F16").value = company1_exit_date

    company2_exit_date = company2_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F17").value = company2_exit_date
    
    company3_exit_date = company3_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F18").value = company3_exit_date
    
    company4_exit_date = company3_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F19").value = company4_exit_date
    
    company5_exit_date = company4_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F20").value = company5_exit_date
    
    company6_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F21").value = company6_exit_date
    
    company7_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F22").value = company7_exit_date
    
    company8_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F23").value = company8_exit_date
    
    company9_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F24").value = company9_exit_date
    
    company10_exit_date = company1_inv_date + timedelta(days = np.random.triangular(hold_period_lower_bound*days,
                                                        hold_period_mode*days,
                                                        hold_period_upper_bound*days))
    model.range("F25").value = company10_exit_date

    #portfolio output
    portfolio_MOIC = model.range("C28").value
    portfolio_IRR = model.range("C29").value
    
    results.append((portfolio_MOIC,portfolio_IRR))
    return results