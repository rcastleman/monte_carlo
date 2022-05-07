import xlwings as xw
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
import matplotlib.pyplot as plt
from datetime import date,timedelta
from pyxirr import xirr

# connect workbook to program
book = xw.Book('simple_mc.xlsx')  

#define which sheet is the Model
model = book.sheets('model')
results = book.sheets('results')
results.clear_contents()

#add a sheet to collect results
# book.sheets.add('results')

#write a value to the input cell in the Model
# model.range('A1').value = 11

#define input distribution
def input():
    input_mean = 15
    input_std = 5
    return random.normalvariate(input_mean,input_std)

#define the input and output cells
# input = model.range('A1').value
# model.range('A1').value = input_random
# input = model.range('A1').value
# output = model.range('A2').value

#simulation loop
input_list = []
output_list = []
num_sims = 10
for i in range(num_sims):
    model.range('A1').value = input()
    output = model.range('A2').value
    input_list.append(model.range('A1').value)
    output_list.append(output)

#transpose inputs & outputs to be exported to worksheet
output_list_tranposed = [[x] for x in output_list]
input_list_transposed = [[x] for x in input_list]

#export results to Results tab of worksheet
results.range('A1').value = 'Inputs'
results.range('B1').value = 'Outputs'
results.range('A2').value = input_list_transposed
results.range('B2').value = output_list_tranposed

# print(output_list)
# print(book.sheets)
# print(f'The square of {input:.1f} is {output:.1f}')