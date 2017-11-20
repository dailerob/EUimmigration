#file for loading data from unfiltered excel files

import numpy as np
import  pandas as pd
import xlwings as xw

def Load_unemployment():
    print("loading unemployment")
    wb = xw.book("dataSets/Unfiltered Unemployment Rates.xlsx")
