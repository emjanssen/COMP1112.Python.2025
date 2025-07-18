import ezsheets
#python library to help with generating 5000 random names
import names

# - - - Function Declarations - - - #

def generateRandomNames():
        
    numberOfNamesRequired = 5000
    listofUniqueNames = []

    while len(listofUniqueNames) < numberOfNamesRequired:
        #call get_full_name() from names library to generate a random full name
        fullName = names.get_full_name()
        #get_full_name() does not have any built-in checks for duplicate, so verify if it's in listofUniqueNames list before adding to list
        if fullName not in listofUniqueNames:
            listofUniqueNames.append(fullName)
    #print(f"These are our unique names: {listofUniqueNames}") // debug statement
    return listofUniqueNames

    #create new spreadsheet, generate sheet with required row/column parameters, delete default sheet that is created upon new spreadsheet generation
def generateSheetStructure():
    week10Spreadsheet = ezsheets.createSpreadsheet(title='Week 10 Python Assignment')
    sheetOne = week10Spreadsheet.createSheet(title="Week10Sheet1", rowCount = 1000, columnCount = 5)
    week10SpreadsheetDefaultSheet = week10Spreadsheet.sheets[0]
    week10SpreadsheetDefaultSheet.delete()
    
def populateSheet():
    return 0

def readSheet():
    # activeSpreadsheet = ezsheets.Spreadsheet('[address]')
    return 0

# - - - Function Calls - - - #

#generateSheetStructure() - this works
listofUniqueNames = generateRandomNames()