from functions import *
from config import *

wk = accessSheet(credentialsFile, accessScope, spreadsheetName)
dictionary = createDictionary(wk, startCol, startRow, endCol)
summarize(wk, dictionary)
