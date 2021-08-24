from Functions import *

credentialsFile = 'Credentials.json'
accessScope = gspread.auth.READONLY_SCOPES
spreadsheetName = "Copy of EnactusHacks 2.0 Participant Feedback (Responses)"

wk = accessSheet(credentialsFile, accessScope, spreadsheetName)

getColValues(wk, 3)