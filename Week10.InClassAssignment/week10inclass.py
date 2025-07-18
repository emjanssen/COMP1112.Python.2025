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
        #for future, will probably learn to use set(), because checking one by one like this takes a long time
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
    return sheetOne

def populateSheet(sheetOne, listofUniqueNames):
    namesPerColumn = 1000
    totalColumns = 5

    #for loop that runs five times
    for column in range(totalColumns):
        #finds starting index by using value of column (i.e. loop count), multiplied by namesPerColumn (i.e. 1000)
        #on loop 1, starting index will be 0 (0 * 1000); on loop 2, starting index will be 1000 (1 * 1000); etc
        startingIndex = column * namesPerColumn
        #the ending index will be whatever the current starting index is, plus 1000
        #so is startingIndex is 1000, endingIndex will be 2000, and so forth
        endingIndex = startingIndex + namesPerColumn
        #split listofUniqueNames using the starting and index indexes on each loop
        #add that split list of names to chunkOfNames variable
        chunkOfNames = listofUniqueNames[startingIndex:endingIndex]

        #on each loop, use the updateColumn() function to update/print the current value of chunkOfNames to the loop's current column
        #we add 1 because google sheets aren't 0 indexed; column A is column 1, B is 2, and so forth
        sheetOne.updateColumn(column + 1, chunkOfNames)

# - - - Function Calls - - - #

sheetOne = generateSheetStructure()
listofUniqueNames = generateRandomNames()
populateSheet(sheetOne, listofUniqueNames)