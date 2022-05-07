from turtle import color
from matplotlib import lines
import xlwings as xw
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
import matplotlib.pyplot as plt
from datetime import date,timedelta
from pyxirr import xirr

#guided by https://www.youtube.com/watch?v=Tv701NoFKw8

#Important note: the calculations are done on the Excel spreadsheet, but the collection of the inputs and outputs is done in the program using dataframes.  For large #s of simluations, this will be slow and the calculation should be moved into the program -- though the summary stats and plots can still be written out to a spreadsheet.

# connect workbook to program
book = xw.Book('simple_mc.xlsx')  

#define which sheet is the Model and which is the Results
model = book.sheets('model')
results = book.sheets('results')
results.clear_contents()


#define the input distribution function (just normal in this case)
def input():
    input_mean = 10 
    input_std = 1.2
    return random.normalvariate(input_mean,input_std)

#simulation 
num_sims = 1000
input_list = []
output_list = []
for i in range(num_sims):
    model.range('A1').value = input() #writes the input value to the appropriate cell so the worksheet can use it 
    output = model.range('A2').value #reads the output value from the worksheet based on the input above
    input_list.append(model.range('A1').value) #adds this particular input value to the list of inputs
    output_list.append(output) #adds this particular output to the list of outputs

#transpose inputs & outputs to be exported to worksheet
output_list_tranposed = [[x] for x in output_list]
input_list_transposed = [[x] for x in input_list]

#export results to Results tab of worksheet
results.range('A1').value = 'Inputs'
results.range('B1').value = 'Outputs'
results.range('A2').value = input_list_transposed
results.range('B2').value = output_list_tranposed

#export results to dataframes
df = pd.DataFrame(output_list)
df.columns = ['Outputs']
print(df.head())

#create plot(s)
sim_fig = plt.figure()
plot = plt.hist(df,
        density=True,
        bins=10)
plt.xlabel('Outputs')
plt.ylabel('Density')
plt.title('Distribution of Outcomes')
plt.vlines(df.mean(),
    ymin = 0,
    ymax = 0.025,
    color='red')

#EXPORT FIGURES TO WORKSHEET

#histogram of results
rng = results.range('G2')
results.pictures.add(sim_fig,
    name = 'Simulation',
    update = True,
    top = rng.top,
    left = rng.left)

#Data series summary stats
description = df.describe()
results.range('D2').value = description

#Cumulative Distribution 
fig_cdf = plt.figure()
x = np.sort(df['Outputs'])
y = np.arange(1,len(x)+1)/len(x)
plt.plot(x,y,
    marker = '.',
    linestyle = None)
plt.xlabel = ('Outputs')
plt.title('Cumulative Distribution Function')
plt.plot(x,y)
plt.show

rng = results.range('G24')
results.pictures.add(fig_cdf,
    name = 'Cumul Dist Function',
    update = True,
    top = rng.top,
    left = rng.left)

# print(plot)
# print(output_list)
# print(book.sheets)
# print(f'The square of {input:.1f} is {output:.1f}')