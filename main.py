from functions import *
from config import *

spreadSheet = accessSheet(credentialsFile, accessScope, spreadsheetName)
dictionary = createDictionary(spreadSheet[1], startCol, startRow, endCol)
summary = summarize(spreadSheet[1], dictionary)
writeSummary(spreadSheet[0], summary, startCol, startRow, endCol, endRow)