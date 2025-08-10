# - - - Imports - - - #

import os
import docx
import subprocess

# - - - Function Declarations - - - #

def createAgendaDirectory():

    agendaTemplateLocationDirectory = r"C:\Users\E.M. Janssen\Documents\GitHub\COMP1112.Python.2025\FinalProject.2025.08.10\Process2.AgendaTemplate"

    if not os.path.exists(agendaTemplateLocationDirectory):
        os.makedirs(agendaTemplateLocationDirectory)
        print("Directory created...")
    else:
        print("Directory already exists...")
    return agendaTemplateLocationDirectory

import os
import subprocess

def openAgendaTemplate(agendaTemplateLocationDirectory):

    templateFileName = "AgendaTemplate.docx"
    templateFile = os.path.join(agendaTemplateLocationDirectory, templateFileName)
    pathToWord = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"

    if not os.path.exists(templateFile):
        print(f"File not found at: {templateFile}")
        return

    try:
        os.startfile(templateFile)
    except FileNotFoundError as exception:
        print(f"File not found: {exception}")
    except Exception as exceptione:
        print(f"An error occurred while opening the file: {exception}")

# - - - Function Calls - - - #

agendaTemplateLocationDirectory = createAgendaDirectory()
openAgendaTemplate(agendaTemplateLocationDirectory)