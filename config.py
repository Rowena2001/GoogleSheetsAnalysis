import gspread

credentialsFile = '/Users/rowenashi/TechProjects/Credentials/Credentials.json'
accessScope = gspread.auth.DEFAULT_SCOPES
spreadsheetName = "Copy of EnactusHacks 2.0 Participant Feedback (Responses)"

startCol = 'C'
startRow  = '1'
endCol = 'H'
endRow = '16'