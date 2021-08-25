from functions import *
from config import *

spreadSheet = accessSheet(credentialsFile, accessScope, spreadsheetName)
dictionary = createDictionary(spreadSheet[1], startCol, startRow, endCol)
summarize(spreadSheet[1], dictionary)