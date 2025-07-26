"""
Name: Elizabeth Janssen
Student ID: 200627709
Course: COMP1112
Assignment: Week 11 - In Class Assignment
Date: 2025.07.25 
Program Description: Prompts the user with a list of words provided via a Notepad file, and time how long it takes the user to type the words.
"""

# - - - To-Do - - - #

#1. Add a function that validates whether the user input into Notepad matches the sentence in the file
#2. Better process for closing file:
    #At present, code relies on user to terminate notepad
    #Better design would maybe be to listen for a keystroke (like the enter key) in Notepad if that's possible? 
    #And then call openedFile.terminate() once we register that specific keystroke 
    #We could ask the user for a specific keystroke in the console once they're ready to close the file
    #But having them go back to the console would extend how long the Notepad file is open
    #So we'd have a less accurate reflection of how long user took to enter the words into the Notepad file

# - - - Imports - - - #

import time
import subprocess

# - - - Function Declarations - - - #

#Function one: wait n number of seconds
#Will use this for any waiting requirements in program
def waitInSeconds(timeToWait):
    time.sleep(timeToWait)

#Function two: wait 30 seconds before starting program
def initialWait():
    timeToWait = 30
    print(f"Waiting {timeToWait} seconds before we start the program...")
    waitInSeconds(timeToWait)
    print(f"Done waiting {timeToWait} seconds.")

#Function three: after user has entered input, wait ten seconds before printing how long it took them to enter the words
def postUserInputWait():
    timeToWait = 10
    print(f"Waiting {timeToWait} seconds before we report how long it took for user to write the words...")
    waitInSeconds(timeToWait)
    print(f"Done waiting {timeToWait} seconds.")

#Function four: provide user with instructions
def provideInstructions():
    #buffer so user has time to read the instructions
    timeToWait = 10
    print("In ten seconds, a Notepad file will open, and it will contain a sentence.\nPlease type that sentence into Notepad exactly as it's written.\nYou will be time on how long it takes you to write the sentence.\nPlease close the file when you're done. You don't need to save the file.")
    print(f"Waiting {timeToWait} seconds...")
    waitInSeconds(timeToWait)
    print("Opening Notepad file...")

#Function five: open Notepad file using Popen
#For this assignment, working off the assumption that the file exists and the path is correct
#In real code, would include safeguards to validate file path, and create file if it doesn't exist yet
def openNotepadFile():
    #raw string so we don't have to escape backslashes
    notepadFilePath = r"C:\Users\E.M. Janssen\Documents\GitHub\COMP1112.Python.2025\Week11.InClassAssignment\listOfWords.txt"
    #save opened file value
    openedFile = subprocess.Popen(["notepad.exe", notepadFilePath])
    return openedFile

#Function six: time how long the Notepad file is open for 
def timeHowLongFileIsOpen(openedFile):
    #get the start time value
    startTime = time.time()
    #start waiting 
    openedFile.wait()
    #get the end time value once the file is closed; i.e. pause script until we exit Notepad
    endTime = time.time()
    #subtract start time from end time for file open duration
    elapsedTime = endTime - startTime
    return elapsedTime

#Function seven: print how long the user took to type the sentence
def printUserInputDuration(elapsedTime):
    print(f"You took {elapsedTime} seconds to type the sentence in Notepad.")

# - - - Function Calls - - - #

initialWait()
provideInstructions()
openedFile = openNotepadFile()
elapsedTime = timeHowLongFileIsOpen(openedFile)
postUserInputWait()
printUserInputDuration(elapsedTime)