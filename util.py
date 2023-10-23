import os
import openpyxl
import skcriteria as skc
from skcriteria.preprocessing import invert_objectives, scalers
from skcriteria.madm import similarity
from skcriteria.pipeline import mkpipe
import datetime

def get_file_direct(filePath, file_name, error_label):
    print(filePath)
    try:
        os.chdir(filePath)
        wb = openpyxl.load_workbook(file_name)
        return wb
    except Exception as error:
        print('Please make sure that you have inputed a valid path. You can get the path by right-clicking the file and selecting "Copy Path".',"\n")
        print(error)
        error_label.config(text = 'Please make sure that you have inputed a valid path.\n You can get the path by right-clicking the file and selecting "Copy Path".')
        return "error"

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

def display_snapshot(sheet, last_row, alternatives, criteria, weights, objectives, matrix):
    #display the alternatives
    for alternative in alternatives:
            sheet.cell(row= last_row + 3 + alternatives.index(alternative), column= 1).value = alternative

    #display the criteria
    for criterion in criteria:
        sheet.cell(row= last_row, column= 2 + criteria.index(criterion)).value = criterion

    #display the weights
    for weight in weights:
        sheet.cell(row= last_row + 1, column= 2 + weights.index(weight)).value = weight

    #display the objectives
    for objective in objectives:
        sheet.cell(row= last_row + 2, column= 2 + objectives.index(objective)).value = objective

    #display matrix
    for row in matrix:
        for value in row:
            sheet.cell(row= last_row + 3 + matrix.index(row), column= 2 + row.index(value)).value = value

def display_rankings(sheet, rank_column, last_row, rankings):
    #inserting rank header with date & time
    sheet.cell(row= last_row, column=rank_column).value = f'Rankings:'
    sheet.cell(row= last_row, column=rank_column).fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=openpyxl.styles.colors.Color(rgb='00FFFF00'))
    #inserting ranks
    for rank in rankings:
        sheet.cell(row= last_row + 3 + rankings.index(rank), column=rank_column).value = rank

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
    