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

# connect workbook to program
book = xw.Book('returns_model.xlsx')  

#define Model and Results sheets
model = book.sheets('Model')
results = book.sheets('Results')

#connect XLS static values -> program
S1_date = model.range("B2").options(dates=dt.date).value
S2_date = model.range("B3").options(dates=dt.date).value
S3_date = model.range("B4").options(dates=dt.date).value
S4_date = model.range("B5").options(dates=dt.date).value
S5_date = model.range("B6").options(dates=dt.date).value

num_series = 5
row_of_first_series = 2 

series_list = []

for i in range(num_series):
    series_name = f"Series {i+1}"

    series_date = model.range(f"B{i+row_of_first_series}").options(dates=dt.date).value #depends on where the cell is
    print(series_date)

    series_duration = model.range(f"C{i+row_of_first_series}").value
    print(series_duration)

    # series = Series(series_name, series_date, series_duration, ...)
    # series_list.append(series)

S1_duration = model.range("C2").value
S2_duration = model.range("C3").value
S3_duration = model.range("C4").value
S4_duration = model.range("C5").value
S5_duration = model.range("C6").value

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
