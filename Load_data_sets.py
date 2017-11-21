#file for loading data from unfiltered excel files

import numpy as np
import pandas as pd
import re
import xlwings as xw

#formats elements from strings into floats
def formatElement(unformated):

    #EuroStat keeps nan data as ":"
    if unformated == ":" or unformated == ": ":
        return np.NaN
    else:
        pulledFloats = re.findall("\\d+(?:\\.\\d+)?", str(unformated))
        if len(pulledFloats) == 0:
            return np.NaN
        return float(pulledFloats[0])

def pullRegion(rowVal):
    return rowVal[rowVal.rfind(',')+1:]

#removes spare characters and formats NANs in EuroData
def filter_unemployment():
    print("loading unemployment")
    wb = xw.Book("dataSets/Unfiltered Unemployment Rates.xlsx")
    currentSheet = wb.sheets['Sheet1']

    df = currentSheet.range('A1').options(pd.DataFrame,expand='table').value

    #formats each element to remove unwanted characters, etc.
    df = df.applymap(formatElement)

    #formats column headers to represent time in floats
    df.columns = map(formatElement, df.columns)

    #format index to get region
    df.index = map(pullRegion, df.index)

    return df





