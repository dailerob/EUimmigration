#file for loading, filtering, and preping data from unfiltered excel files

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

#pulls the region code for element
def pullRegion(rowVal):
    return rowVal[rowVal.rfind(',')+1:]

#removes spare characters and formats NANs in EuroData
def filter_dataset(setName):
    print("loading dataSet")
    wb = xw.Book(setName)
    currentSheet = wb.sheets['Sheet1']

    df = currentSheet.range('A1').options(pd.DataFrame,expand='table').value

    #formats each element to remove unwanted characters, etc.
    df = df.applymap(formatElement)

    #formats column headers to represent time in floats
    df.columns = map(formatElement, df.columns)

    #format index to get region
    df.index = map(pullRegion, df.index)

    return df

#isolates statistics for just the NUTS-2 regions
def isolate_Nuts_2(df):
    removeList = []

    for i in range(df.shape[0]):
        floatsFound = re.findall("\\d+(?:\\.\\d+)?", str(df.index[i]))
        if len(floatsFound)==0:
            currentIndex = 0
        else:
            currentIndex = float(floatsFound[0])

        if currentIndex < 10 or currentIndex >99:
            removeList.append(i)
    df = df.drop(df.index[removeList])

    return df

#sets the nans to the regional average
def removeNaNs(df):
    for i in range(df.shape[0]):
        rowAverage = np.mean(df.iloc[i].dropna())
        df.iloc[i] = df.iloc[i].fillna(rowAverage)

    return df

def removeFixedEffects(df):
    for i in range(df.shape[0]):
        rowAverage = np.mean(df.iloc[i])
        df.iloc[i] = df.iloc[i] - rowAverage

    for i in range(len(df.columns)):
        columnAverage = np.mean(df[df.columns[i]])

        df[df.columns[i]] = df[df.columns[i]] - columnAverage

    return df


#reads in unemployment and migration data, filters them, and then
#lines them up by region
def prepForModel():

    unemploy = filter_dataset("dataSets/Unfiltered Unemployment Rates.xlsx")
    unemploy = isolate_Nuts_2(unemploy)
    unemploy = removeNaNs(unemploy)
    unemploy = removeFixedEffects(unemploy)

    migration = filter_dataset("dataSets/Net Migration Rates.xlsx")
    migration = isolate_Nuts_2(migration)
    migration = removeNaNs(migration)
    migration = removeFixedEffects(migration)

    migration = migration.drop(2016.0,axis=1)
    unemploy = unemploy.drop(2016.0, axis=1)
    unemploy = unemploy.drop(1999.0, axis=1)

    xlist = []
    ylist = []

    for i in range(unemploy.shape[0]):
        region = unemploy.index[i]
        if region in migration.index:
            migVals = migration.loc[region].values.tolist()
            unVals = unemploy.loc[region].values.tolist()

            #sanity check to make sure our data lines up
            if not len(migVals) == len(unVals):
                print(str(unemploy.index[i]))
                print(str(len(migVals)))
                print(str(len(unVals)))

            xlist.extend(migVals)
            ylist.extend(unVals)

    return (xlist,ylist)











