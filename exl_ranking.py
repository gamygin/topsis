import os
import openpyxl
import sys


if __name__ == "__main__":
    # Asks the user to enter the filepath of the excel file.
    filePath = input('Please enter the path of the folder where the excel files are stored: ')
    # Goes inside that folder.
    os.chdir(filePath)
    #gets excel file
    file_name = input('Please enter the name of the file: ')
    #loads excel file
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    #main task
    while True:
        matrix
        break