from matplotlib import lines
import xlwings as xw
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
from turtle import color
import matplotlib.pyplot as plt
import datetime as dt
from pyxirr import irr, xirr

"""Strategy:
1. connect program to excel
2. read input variables from excel
3. read distribution parameters from excel
4. define Simulation class
    4(a) define Series class?
5. run [x] simulations
6. collect data from each simulation
7. send data to excel
8. send distribution plots to excel
"""

# connect workbook to program
book = xw.Book('returns_model.xlsx')  

#define Model and Results sheets
model = book.sheets('Model')
results = book.sheets('Results')
results.clear_contents()

#read distribution parameters
duration_min = model.range("J14").value
duration_mode = model.range("K14").value
duration_max = model.range("L14").value
series_stepup_mean = model.range("K15").value
series_stepup_sd = model.range("M15").value
exit_stepup_mean = model.range("K16").value
exit_stepup_sd = model.range("M16").value

num_sims = 100
simulation_results = []

for sim in range(num_sims):
    #use np.triangular to generate duration
    model.range("C2").value = np.random.triangular(duration_min, duration_mode, duration_max)
    model.range("C3").value = np.random.triangular(duration_min, duration_mode, duration_max)
    model.range("C4").value = np.random.triangular(duration_min, duration_mode, duration_max)
    model.range("C5").value = np.random.triangular(duration_min, duration_mode, duration_max)
    model.range("C6").value = np.random.triangular(duration_min, duration_mode, duration_max)

    #use np.normal to generate stepup values
    model.range("E3").value = np.random.normal(series_stepup_mean, series_stepup_sd)
    model.range("E4").value = np.random.normal(series_stepup_mean, series_stepup_sd)
    model.range("E5").value = np.random.normal(series_stepup_mean, series_stepup_sd)
    model.range("E6").value = np.random.normal(series_stepup_mean, series_stepup_sd)
    
    #use np.normal to geneate exit stepup value 
    model.range("C10").value = np.random.lognormal(exit_stepup_mean, exit_stepup_sd)
   
    #read exit results
    exit_date = model.range("B9").value
    exit_stepup = model.range("C10").value
    exit_proceeds = model.range("C9").value
    MOIC = model.range("B13").value
    IRR = model.range("B14").value

    simulation_results.append([MOIC,IRR]) 

#collect results in dataframe
output_data = pd.DataFrame(simulation_results,columns = ["MOIC","IRR"])
output_data.index.name = 'Sim #'

#output data -> Results tab
results.range('A1').value = output_data

#CREATE PLOTS

#probability density function
MOIC_PDF_fig = plt.figure()
plot = plt.hist(output_data["MOIC"],
        density=True,
        bins=10)
plt.xlabel('Outputs')
plt.ylabel('Density')
plt.title('MOIC - Probability Distribution')
plt.vlines(output_data["MOIC"].mean(),
    ymin = 0,
    ymax = 0.20,
    color='red')

IRR_PDF_fig = plt.figure()
plot = plt.hist(output_data["IRR"],
        density=True,
        bins=10)
plt.xlabel('Outputs')
plt.ylabel('Density')
plt.title('IRR - Probability Distribution')
plt.vlines(output_data["IRR"].mean(),
    ymin = 0,
    ymax = 2.5,
    color='red')

#cumulative distribution function
MOIC_CDF_fig = plt.figure()
x = np.sort(output_data["MOIC"])
y = np.arange(1,len(x)+1)/len(x)
plt.plot(x,y,
    marker = '.',
    linestyle = None)
plt.xlabel = ('Outputs')
plt.title('MOIC - Cumulative Distribution Function')
plt.plot(x,y)
plt.show

IRR_CDF_fig = plt.figure()
x = np.sort(output_data["IRR"])
y = np.arange(1,len(x)+1)/len(x)
plt.plot(x,y,
    marker = '.',
    linestyle = None)
plt.xlabel = ('Outputs')
plt.title('IRR - Cumulative Distribution Function')
plt.plot(x,y)
plt.show

#dataframe description
description = output_data.describe()

#PLOTS -> Results tab
results.range('E1').value = description

rng_1 = results.range('I1')
results.pictures.add(MOIC_PDF_fig,
    name = 'MOIC PDF',
    update = True,
    top = rng_1.top,
    left = rng_1.left)

rng_2 = results.range('I29')
results.pictures.add(IRR_PDF_fig,
    name = 'IRR PDF',
    update = True,
    top = rng_2.top,
    left = rng_2.left)

rng_3 = results.range('P1')
results.pictures.add(MOIC_CDF_fig,
    name = 'MOIC CDF',
    update = True,
    top = rng_3.top,
    left = rng_3.left)

rng_4 = results.range('P29')
results.pictures.add(IRR_CDF_fig,
    name = 'IRR CDF',
    update = True,
    top = rng_4.top,
    left = rng_4.left)


#DISCARDS
# class Series:
#     def __init__(self, name, date, duration, target_percent, total_capital):
#         self.name = name
#         self.date = date
#         self.duration = duration
#         self.target_percent = target_percent
#         self.total_capital = total_capital

#define number of series in case that changes someday(?)
# num_series = 5
#define positional offsets to target correct cells
# first_row = 2 
# initial_pre_money = model.range("F2").value

# for i in range(num_series):  #defining in terms of num_series may not be the best way to do this... 
#     series_name = model.range(f"A{i + first_row}").value
#     print("\r\n")
#     print(f"Series Name: {series_name}")
#     series_date = model.range(f"B{i + first_row}").options(dates=dt.date).value #depends on where the cell is
#     print(f"Series Date: {series_date}")
#     series_duration = model.range(f"C{i + first_row}").value
#     print(f"Series Duration: {series_duration} days")
#     series_target_percent = model.range(f"D{i + first_row}").value
#     print(f"Investor Target %: {series_target_percent*100}")
#     series_stepup = model.range(f"E{i + first_row}").value
#     print(f"Step-up from Last Series: {series_stepup}x")
#     series_total_capital = model.range(f"F{i + first_row}").value
#     print(f"Series Total Capital: ${series_total_capital}")
#     series_investor_capital = series_total_capital * series_target_percent
#     print(f"Investor Capital: ${series_investor_capital}")

    # series = Series(series_name, series_date, series_duration, ...)
    # series_list.append(series)

# for i in range(1,6):
#     temp_date = globals()[f"S{i}_date"]
#     temp_dur = globals()[f"S{i}_duration"]
#     print(temp_date,temp_dur)

# s1_target = model.range("H17").value
# s2_target = model.range("H18").value
# s3_target = model.range("H19").value
# s4_target = model.range("H20").value
# s5_target = model.range("H21").value
# s1_pre = model.range("H22").value
# s1_total_capital = model.range("H23").value
# s2_total_capital = model.range("H24").value
# s3_total_capital = model.range("H25").value
# s4_total_capital = model.range("H26").value
# s5_total_capital = model.range("H27").value