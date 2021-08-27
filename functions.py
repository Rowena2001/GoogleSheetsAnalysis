import gspread
from statistics import mean, median, mode
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from wordcloud import WordCloud
import matplotlib.pyplot as plt

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

    return [sh, wk]

# creates a dictionary of all the values inside the Google spreadsheet
# uses functions getRowValues and getColValues
# returns a dictionary
def createDictionary(wk, startCol, startRow, endCol):

    headerList = getHeaders(wk, startCol, startRow, endCol)
    valuesList = getColValues(wk, startCol, endCol)

    #created empty dictionary
    spreadsheetValuesDict = {}

    # populates dictionary using headers as the keys
    for i in range(len(headerList[0])):
        spreadsheetValuesDict[headerList[0][i]] = valuesList[i]

    return spreadsheetValuesDict

# determines summarizing method depending on data type
def summarize(wk, dictionary):

    print("\nSUMMARY\n")

    summary = {}

    for key in dictionary.keys():
        print("\n", key, "\n")
        frequencyDictionary = createFrequencyDictionary(wk, dictionary[key])
        if numericAnswer(frequencyDictionary):
            summary[key] = computeAverages(dictionary[key])
        elif multipleChoice(frequencyDictionary):
            summary[key] = (formatMCSummary(frequencyDictionary))
        else:
            createWordCloud(frequencyDictionary)

    return summary

def writeSummary(sh, summary, startCol, startRow, endCol, endRow):
    try:
        summarySheet = sh.add_worksheet("Summary", 100, 100)
    except:
        summarySheet = sh.worksheet("Summary")
        summarySheet.clear()
    
    row = 1
    col = 1
    for key in summary.keys():
        summarySheet.update_cell(row, col, key)
        row += 1
        for values in summary[key]:
            indentCol = 2
            for value in values:
                summarySheet.update_cell(row, indentCol, value)
                indentCol += 1
            row += 1

# gets all values inside row of worksheet wk
# returns a list of all the values
def getHeaders(wk, startCol, startRow, endCol):

    # gets header values
    valuesList = wk.get(getStringRange(startCol, startRow, endCol, startRow))

    return valuesList

# gets all values inside column col of worksheet wk
# removes the first value as this represents the header
# returns a list of all the values
def getColValues(wk, startCol, endCol):

    valuesList = []

    numericStartCol = ord(startCol) - 64
    numericEndCol = ord(endCol) - 64

    for i in range(numericEndCol - numericStartCol + 1):

        colValuesList = wk.col_values(numericStartCol + i)
        del colValuesList[0]

        valuesList.append(colValuesList)

    return valuesList

# tracks the frequency of each answer
def createFrequencyDictionary(wk, values):
    # creates empty dictionary
    frequencyDictionary = {}

    # populates dictionary with answers as keys, and frequency as values
    for value in values:
        answersList  = value.split(", ")
        for answer in answersList:
            # if answer exists as a key, add 1 to the frequency
            if answer in frequencyDictionary.keys():
                frequencyDictionary[answer] += 1
            # answer does not exist as a key, make frequency 1
            else:
                frequencyDictionary[answer] = 1

    return frequencyDictionary

# checks if the keys/answers of a dicationary are numeric
# returns true if numeric
def numericAnswer(dictionary):
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

# checks if the answers are multiple choice by seeing if the frequencies are greater than  1
# this assumes that non-MC answers are all different, meaning that the frequency must be 1
# returns true if MC
def multipleChoice(dictionary):
    isMC = False
    for frequency in dictionary.values():
        if frequency > 1:
            isMC = True
            break
    return isMC

# computes averages and returns them in a list
# returns a list of all the values
def computeAverages(values):
    meanValue = mean(int(x) for x in values)
    medianValue = median(int(x) for x in values)
    modeValue = mode(int(x) for x in values)

    print("\tMean: ", meanValue, "\n\tMedian: ", medianValue, "\n\tMode: ", modeValue, "\n")
    return [["Mean", str(meanValue)], ["Median", str(medianValue)], ["Mode", str(modeValue)]]

# formats and prints the MC summary
def formatMCSummary(dictionary):
    MCSummary = []
    for key in dictionary.keys():
        print("\t", key, ":", dictionary[key])
        MCSummary.append([key, dictionary[key]])
    return MCSummary

# creates a word cloud for non-numeric and non-MC answers
# visually shows the most common words
def createWordCloud(dictionary):
    words = []
    stopWords = set(stopwords.words("english"))

    for key in dictionary.keys():
        keySplit = key.split()
        for word in keySplit:
            if word not in stopWords:
                words.append(word.strip('''!()-[]{};:'"\,<>./?@#$%^&*_~''').lower())
    
    joinWords = ' '.join([word for word in words])

    wordcloud = WordCloud(width=1600, height=800, max_font_size=200, background_color="black").generate(joinWords)
    plt.figure(figsize=(12,10))
    plt.imshow(wordcloud)
    plt.axis("off")
    # plt.show()
    return wordcloud

def getStringRange(startCol, startRow, endCol, endRow):
    #  sets header cell range
    startCell = startCol + startRow
    endCell = endCol + endRow
    return startCell+':'+endCell