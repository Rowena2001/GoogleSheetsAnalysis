from functions import *
from config import *

wk = accessSheet(credentialsFile, accessScope, spreadsheetName)
getColValues(wk, startCol, startRow, endRow)
# createDictionary(wk, startCol, startRow, endCol, endRow)