import os
import openpyxl
import skcriteria as skc
from skcriteria.preprocessing import invert_objectives, scalers
from skcriteria.madm import similarity
from skcriteria.pipeline import mkpipe
import datetime

def get_file_path():
    # Asks the user to enter the filepath of the excel file.
    filePath = input('Please enter the path of the folder where the excel file is stored: ')
    # Goes inside that folder.
    try:
        os.chdir(filePath)
        file = get_file()
        return file
    except:
        print("Input a valid path. Do not input the path to the file itself but rather to its location.","\n")
        get_file_path()

def get_file():
    #gets excel file
    file_name = input('Please enter the name of the file: ')
    #loads excel file
    try:
        wb = openpyxl.load_workbook(file_name)
        return wb, file_name
    except:
        print('Input a valid file. Check for spelling mistakes and ".xlsx". You might also be in the wrong directory',"\n")
        get_file()

def create_matrix(sheet, row, column):
    matrix = []
    for r in range(row-4):
        values = []
        for c in range(column-2):
            try:
                value = float(sheet.cell(row=r+4, column=c+2).value)
            except:
                print('Incorrect values',"\n")
                break
            values.append(value)
        matrix.append(values)
    print(matrix,"\n")
    return matrix

def column_index_to_letters(rank_column):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    if rank_column <= len(alphabet):
        index = alphabet[rank_column-1]
    else:
        t = rank_column//len(alphabet) 
        for i in range(t):
            index = "Z"*t
        index += alphabet[rank_column-len(alphabet*t)-1]

    return index
    