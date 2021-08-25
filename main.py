from functions import *
from config import *

wk = accessSheet(credentialsFile, accessScope, spreadsheetName)

getColValues(wk, startCol, endCol)
createDictionary(wk, startCol, startRow, endCol)
