# - - - Imports - - - #

import os
from docx import Document

# - - - Template Inputs - - - #

#admin edits these variable values based on what they'd like their new modified agenda to say
newTitle        = "Council of Elrond"
newDate         = "2025/08/16"
newTime         = "13:30"
newLocation     = "A202"
newHost         = "Elrond"
newAttendees    = "Gimli, Boromir"
newAbsents      = "Celeborn, Thranduil"
newDelegates    = "Legolas Greenleaf"
newActionItems  = "1. Review Events To-Date\n2.Explore Options"
newPurpose      = "Discuss growing threat in Mordor."

# - - - Function Declarations - - - #

# Create Directory #

def createAgendaDirectory():

    agendaTemplateLocationDirectory = r"C:\Users\E.M. Janssen\Documents\GitHub\COMP1112.Python.2025\FinalProject.2025.08.10\Process2.AgendaTemplate"

    if not os.path.exists(agendaTemplateLocationDirectory):
        os.makedirs(agendaTemplateLocationDirectory)
        print("Directory created...")
    else:
        print("Directory already exists...")
    return agendaTemplateLocationDirectory

# Open Agenda Template #

def openAgendaTemplate(agendaTemplateLocationDirectory):

    templateFileName = "AgendaTemplate.docx"
    templateFilePath = os.path.join(agendaTemplateLocationDirectory, templateFileName)

    if not os.path.exists(templateFilePath):
        print(f"File not found at: {templateFilePath}")
        return None

    try:
        agendaDocument = Document(templateFilePath)
    except Exception as exception:
        print(f"Error loading document... {exception}")
        return None

    return agendaDocument

# Repalce Placeholder Text #

def replacePlaceholder(agendaDocument, placeholderText, newValue):

    replacedAnyPlaceholders = False

    for table in agendaDocument.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholderText in cell.text:
                    print(f"Found '{placeholderText}' in table cell: {cell.text.strip()}")
                    cell.text = newValue
                    print(f"Replaced with: {newValue}")
                    replacedAnyPlaceholders = True

    if not replacedAnyPlaceholders:
        print(f"We did not find the placeholder: {placeholderText}.")

    return agendaDocument

# Save Modified Agenda #

def saveModifiedAgenda(agendaDocument, agendaTemplateLocationDirectory):
    pathToModifiedAgenda = os.path.join(agendaTemplateLocationDirectory, "ModifiedAgenda.docx")
    agendaDocument.save(pathToModifiedAgenda)
    print(f"Modified agenda saved to: {pathToModifiedAgenda}")

# - - - Function Calls - - - #

agendaTemplateLocationDirectory = createAgendaDirectory()
agendaDocument = openAgendaTemplate(agendaTemplateLocationDirectory)

# A check in case the agenda hasn't loaded properly/has a value of none
if agendaDocument is not None:

    # Call function for each placeholder; pass in the placeholder text, and the new text value for that placeholder
    # Each time function is called, it loops through the table, looking for placeholder text that matches
    agendaDocument = replacePlaceholder(agendaDocument, "titlePlaceholder", newTitle)
    agendaDocument = replacePlaceholder(agendaDocument, "datePlaceholder", newDate)
    agendaDocument = replacePlaceholder(agendaDocument, "timePlaceholder", newTime)
    agendaDocument = replacePlaceholder(agendaDocument, "locationPlaceholder", newLocation)
    agendaDocument = replacePlaceholder(agendaDocument, "hostPlaceholder", newHost)
    agendaDocument = replacePlaceholder(agendaDocument, "attendeesPlaceholder", newAttendees)
    agendaDocument = replacePlaceholder(agendaDocument, "absentsPlaceholder", newAbsents)
    agendaDocument = replacePlaceholder(agendaDocument, "delegatesPlaceholder", newDelegates)
    agendaDocument = replacePlaceholder(agendaDocument, "actionItemsPlaceholder", newActionItems)
    agendaDocument = replacePlaceholder(agendaDocument, "purposePlaceholder", newPurpose)

    saveModifiedAgenda(agendaDocument, agendaTemplateLocationDirectory)

else:
    print("Failed to load our agenda template.")