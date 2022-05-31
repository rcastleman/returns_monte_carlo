from matplotlib import lines
import xlwings as xw
import pandas as pd
import random
import numpy as np
from numpy.random import uniform
from turtle import color
import matplotlib.pyplot as plt
import datetime as dt
from pyxirr import xirr


"""Strategy:
1. connect program to excel
2. read input variables from excel
3. read distribution parameters from excel
4. define Simulation class
5. run [x] simulations
6. collect data from each simulation
7. port data to excel
8. port plots to excel

"""


# connect workbook to program
book = xw.Book('returns_model.xlsx')  

#define Model and Results sheets
model = book.sheets('Model')
results = book.sheets('Results')

#define number of series in case that changes someday(?)
num_series = 5
#define positional offsets to target correct cells
first_row = 2 
initial_pre_money = model.range("F2").value

series_list = []

for i in range(num_series):  #defining in terms of num_series may not be the best way to do this... 
    series_name = model.range(f"A{i + first_row}").value
    print("\r\n")
    print(f"Series Name: {series_name}")
    series_date = model.range(f"B{i + first_row}").options(dates=dt.date).value #depends on where the cell is
    print(f"Series Date: {series_date}")
    series_duration = model.range(f"C{i + first_row}").value
    print(f"Series Duration: {series_duration} days")
    series_target_percent = model.range(f"D{i + first_row}").value
    print(f"Investor Target %: {series_target_percent*100}")
    series_stepup = model.range(f"E{i + first_row}").value
    print(f"Step-up from Last Series: {series_stepup}x")
    series_total_capital = model.range(f"F{i + first_row}").value
    print(f"Series Total Capital: ${series_total_capital}")
    series_investor_capital = series_total_capital * series_target_percent
    print(f"Investor Capital: ${series_investor_capital}")
    # series_post = 

    # series = Series(series_name, series_date, series_duration, ...)
    # series_list.append(series)

# for i in range(1,6):
#     temp_date = globals()[f"S{i}_date"]
#     temp_dur = globals()[f"S{i}_duration"]
#     print(temp_date,temp_dur)
 
class Simulation:
    class Series:...


class Series:
    def __init__(self, name, date, duration, target_percent, total_capital):
        self.name = name
        self.date = date
        self.duration = duration
        self.target_percent = target_percent
        self.total_capital = total_capital


#connect XLS dynamic variable params -> program

#set distributions for dynamic variables

#connection XLS IRR output -> program

#Simulation
num_sims = 500
input_list = []
output_list = []

#collect results in dataframe

#transpose inputs & outputs to be exported to worksheet

#create plot(s)

#Results and Plots -> XLS Results

# print("no bugs currently")
