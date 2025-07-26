"""
Name: Elizabeth Janssen
Student ID: 200627709
Course: COMP1112
Assignment: Week 11 - In Class Assignment
Date: 2025.07.25 
Program Description: Prompts the user with a list of words provided via a Notepad file, and time how long it takes the user to type the words.
"""

# - - - Imports - - - #

import time
import re
import os
import subprocess

# - - - Function Declarations - - - #

#Function one: wait n number of seconds
def waitInSeconds(timeToWait):
    print(f"Waiting {timeToWait} seconds...")
    time.sleep(timeToWait)
    print(f"Done waiting {timeToWait} seconds.")

#Function two: wait 30 seconds before starting program
def initialWait():

    timeToWait = 30
    waitInSeconds(timeToWait)

def provideInstructions():

    timeToWait = 5
    #buffer so user has time to read the instructions
    print("In five seconds, a Notepad fill open, and it will contain a sentence.\nPlease type that sentence into Notepad exactly as it's written.\nYou will be time on how long it takes you to rewrite the sentence.")
    waitInSeconds(timeToWait)
    print("Opening Notepad file...")

#Function three: open Notepad file
def openNotepadFile():

    notepadFilePath = r"C:\Users\E.M. Janssen\Documents\GitHub\COMP1112.Python.2025\Week11.InClassAssignment\listOfWords.txt"
    openedFile = subprocess.Popen(["notepad.exe", notepadFilePath])
    return openedFile

def timeHowLongFileIsOpen(openedFile):
    startTime = time.time()
    openedFile.wait()
    endTime = time.time()
    elapsedTime = endTime - startTime
    print(f"You took {elapsedTime} seconds to type the sentence in Notepad.")
    return elapsedTime

# - - - Function Calls - - - #

#initialWait() - this works; don't call it every time while in dev
provideInstructions()
openedFile = openNotepadFile()
elapsedTime = timeHowLongFileIsOpen(openedFile)