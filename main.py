from Functions import *

__author__ = "Rowena Shi"

credentialsFile = '/Users/rowenashi/TechProjects/Credentials/Credentials.json'
accessScope = gspread.auth.READONLY_SCOPES
spreadsheetName = "Copy of EnactusHacks 2.0 Participant Feedback (Responses)"

wk = accessSheet(credentialsFile, accessScope, spreadsheetName)

spreadsheetValuesDict = createDictionary(wk)

createTallyDictionary(wk, spreadsheetValuesDict["Overall how would you rate EnactusHacks 2.0?"])
