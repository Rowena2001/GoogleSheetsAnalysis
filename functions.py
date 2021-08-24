from typing import SupportsComplex
import gspread
from statistics import *

__author__ = "Rowena Shi"

# accesses the Google spreadsheet and worksheet
# returns the worksheet
def accessSheet(credentialsFile, accessScope, spreadsheetName):
    # Access Credentials.json and create READONLY scopes
    gc = gspread.service_account(filename=credentialsFile, scopes=accessScope)

    # Opens "Copy of EnactusHacks 2.0 Participant Feedback (Responses)" spreadsheet
    sh = gc.open(spreadsheetName)

    # Gets the first sheet
    wk = sh.sheet1

    return wk

# creates a dictionary of all the values inside the Google spreadsheet
# uses functions getRowValues and getcolValues
# returns a dictionary
def createDictionary(wk):
    print("\nHEADERS\n")
    headerList = getRowValues(wk, 1)
    print(headerList)

    spreadsheetValuesDict = {}
    spreadsheetValuesDict = {headerList[i]: getColValues(wk, i+1) for i in range(len(headerList))}

    print("\n\n DICTIONARY VALUES \n\n" + str(spreadsheetValuesDict) + "\n")

    return spreadsheetValuesDict

# gets all values inside row of worksheet wk
# returns a list of all the values
def getRowValues(wk, row):
    # gets row values
    valuesList = wk.row_values(row)

    return valuesList

# gets all values inside column col of worksheet wk
# removes the first value as this represents the header
# returns a list of all the values
def getColValues(wk, col):
    # gets column values
    valuesList = wk.col_values(col)
    del valuesList[0]

    return valuesList

# determines how to analyze metrics depending on type of data
def determineAnalysisType(wk, values):
    pass

# tracks the quantity of each answer
def createTallyDictionary(wk, values):
    tallyDictionary = {}
    for i in values:
        if i in tallyDictionary.keys():
            tallyDictionary[i] += 1
        else:
            tallyDictionary[i] = 1

    print("\n\n TALLY DICTIONARY VALUES \n\n" + str(tallyDictionary) + "\n")
    return tallyDictionary

# computes averages and returns them in a list
# returns a list of all the values
def computeAverages(values):
    meanValue = mean(int(x) for x in values)
    medianValue = median(int(x) for x in values)
    modeValue = mode(int(x) for x in values)
    avgList = [meanValue, medianValue, modeValue]

    return avgList
