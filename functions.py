from typing import SupportsComplex
import gspread
from statistics import *
from pandas import *

__author__ = "Rowena Shi"

# accesses the Google spreadsheet and worksheet
# returns the worksheet
def accessSheet(credentialsFile, accessScope, spreadsheetName):
    # Access Credentials.json and create READONLY scopes
    gc = gspread.service_account(filename=credentialsFile, scopes=accessScope)

    # Opens spreadsheet
    sh = gc.open(spreadsheetName)

    # Gets the first sheet
    wk = sh.sheet1

    return wk

# creates a dictionary of all the values inside the Google spreadsheet
# uses functions getRowValues and getColValues
# returns a dictionary
def createDictionary(wk, startCol, startRow, endCol, endRow):

    headerList = getHeaders(wk, startCol, startRow, endCol,)

    #created empty dictionary
    spreadsheetValuesDict = {}

    # populates dictionary using headers as the keys
    for key in headerList:
        startColLetter = startCol
        spreadsheetValuesDict[key] = getColValues(wk, startColLetter, startRow, endRow)
        startColLetter = chr((ord(startColLetter.upper())+1 - 65) % 26 + 65)

    print("\n\n DICTIONARY VALUES \n\n" + str(spreadsheetValuesDict) + "\n")

    return spreadsheetValuesDict

# determines how to analyze metrics depending on type of data
def analyze(wk, dictionary):
    for key in dictionary.keys():
        frequencyDictionary = createFrequencyDictionary(wk, dictionary[key])
        if numericKeys(frequencyDictionary):
            print(key)
            computeAverages(dictionary[key])

# gets all values inside row of worksheet wk
# returns a list of all the values
def getHeaders(wk, startCol, startRow, endCol):

    print("\nHEADERS\n")

    #  sets header cell range
    startCell = startCol + startRow
    endCell = endCol + startRow
    
    # gets header values
    valuesList = wk.get(startCell+':'+endCell)
    
    print(valuesList)

    return valuesList

# gets all values inside column col of worksheet wk
# removes the first value as this represents the header
# returns a list of all the values
def getColValues(wk, startCol, startRow, endCol, endRow):

    valuesList = []

    # gets column values
    print("COLUMN " + col)

    for i in  range(endCol - startCol):

        # sets column range
        startValueRow = int(startRow) + 1
        startCell = col + str(startValueRow)
        endCell = endCol + endRow

        colValuesList = wk.get(startCell+':'+endCell)
        valuesList.append(colValuesList)

    print(valuesList)

    return valuesList

# converts columns from letter form to numeric form
def getColumnNumbers(startCol, endCol):
    numericStartCol = ord(startCol)
    numericEndCol = ord(endCol)
    print
    return  [numericStartCol, numericEndCol]

# tracks the frequency of each answer
def createFrequencyDictionary(wk, values):
    # creates empty dictionary
    frequencyDictionary = {}

    # populates dictionary with answers as keys, and frequency as values
    for i in values:
        # if answer exists as a key, add 1 to the frequency
        if i in frequencyDictionary.keys():
            frequencyDictionary[i] += 1
        # answer does not exist as a key, make frequency 1
        else:
            frequencyDictionary[i] = 1

    print("\n\n TALLY DICTIONARY VALUES \n\n" + str(frequencyDictionary) + "\n")
    return frequencyDictionary

# checks if the keys of a dicationary are numeric
# returns true if numeric
def numericKeys(dictionary):
    isNumeric = False
    for i in dictionary.keys():
        # sets isNumeric if key is numeric
        # continues iterating through keys
        if i.isdigit():
            isNumeric = True
        # breaks if key is not numeric
        else:
            break
    return isNumeric

# computes averages and returns them in a list
# returns a list of all the values
def computeAverages(values):
    meanValue = mean(int(x) for x in values)
    medianValue = median(int(x) for x in values)
    modeValue = mode(int(x) for x in values)
    avgList = [meanValue, medianValue, modeValue]

    return avgList
