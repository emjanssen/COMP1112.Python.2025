# - - - Create Initial Directory - - - #

from createWorkbook import (
    createDirectory,
    createAdminWorkbook,
    getActiveSheet,
    renameSheet,
    assignColumnHeadings,
    addSecondSheet,
    saveWorkbook,
    returnSheetNames
)

adminWorkbookDirectory = createDirectory()
adminWorkbook = createAdminWorkbook(adminWorkbookDirectory)
sheetOne = getActiveSheet(adminWorkbook)
sheetOne = renameSheet(sheetOne)
assignColumnHeadings(sheetOne)
adminWorkbook = addSecondSheet(adminWorkbook)
saveWorkbook(adminWorkbook, adminWorkbookDirectory)
returnSheetNames(adminWorkbook)