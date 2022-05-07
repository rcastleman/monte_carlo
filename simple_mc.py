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

#write a value to the input cell in the Model
model.range('A1').value = 8

#define the input and output cells
input = model.range('A1').value
output = model.range('A2').value

# print(book.sheets)
print(f'the square of {input} is {output}.')