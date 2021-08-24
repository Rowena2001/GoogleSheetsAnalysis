from typing import SupportsComplex
import gspread

# accesses the Google spreadsheet
def accessSheet(credentialsFile, accessScope, spreadsheetName):
    # Access Credentials.json and create READONLY scopes
    gc = gspread.service_account(filename=credentialsFile, scopes=accessScope)

    # Opens "Copy of EnactusHacks 2.0 Participant Feedback (Responses)" spreadsheet
    sh = gc.open(spreadsheetName)

    # Gets the first sheet
    wk = sh.sheet1

    return wk

def getColValues(wk, col):
    # gets column values
    values_list = wk.col_values(col)

    # prints each value
    for i in values_list:
        print(i)

