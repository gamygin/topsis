import os
import openpyxl
import skcriteria as skc
from skcriteria.preprocessing import invert_objectives, scalers
from skcriteria.madm import similarity
from skcriteria.pipeline import mkpipe

def get_file_path():
    # Asks the user to enter the filepath of the excel file.
    filePath = input('Please enter the path of the folder where the excel file is stored: ')
    # Goes inside that folder.
    try:
        os.chdir(filePath)
        get_file()
    except:
        print("Input a valid path. Do not input the path to the file itself but rather to its location.")
        get_file_path()

def get_file():
    global file_name, wb
    #gets excel file
    file_name = input('Please enter the name of the file: ')
    #loads excel file
    try:
        wb = openpyxl.load_workbook(file_name)
    except:
        print('Input a valid file. Check for spelling mistakes and ".xlsx". You might also be in the wrong directory')
        get_file()



if __name__ == "__main__":
    get_file_path()
    sheet = wb.active
    #getting ready to get info except matrix
    criteria = []
    weights = []
    objectives = []
    alternatives = []
    column = 2
    row = 4
    #getting all the values except matrix
    while True:
        criteriaValue = sheet.cell(row= 1, column= column).value
        criteria.append(criteriaValue)
        weightValue = sheet.cell(row= 2, column= column).value
        try:
            weights.append(float(weightValue))
        except:
            print('Incorrect weights')
            break
        objectiveValue = sheet.cell(row= 3, column= column).value
        if objectiveValue == 'min':
            objectives.append(min)
        elif objectiveValue == 'max':
            objectives.append(max)
        column += 1
        if not sheet.cell(row= 1, column= column).value and not sheet.cell(row= 2, column= column).value and not sheet.cell(row= 3, column= column).value:
            print(criteria, weights, objectives)
            #getting compared object names (alternatives)
            while True:
                alternativeValue = sheet.cell(row= row, column= 1).value
                alternatives.append(alternativeValue)
                row += 1
                if not sheet.cell(row= row, column= 1).value:
                    print(alternatives)
                    break
            break

    #getting matrix
    matrix = []
    for r in range(row-4):
        values = []
        for c in range(column-2):
            try:
                value = float(sheet.cell(row=r+4, column=c+2).value)
            except:
                print('Incorrect values')
                break
            values.append(value)
        matrix.append(values)
    print(matrix)

    #evaluating values
    dm = skc.mkdm(matrix, objectives, weights, alternatives, criteria)
    print(dm,"\n")

    pipe = mkpipe(
    invert_objectives.NegateMinimize(),
    scalers.VectorScaler(target="matrix"),  # this scaler transform the matrix
    scalers.SumScaler(target="weights"),  # and this transform the weights
    similarity.TOPSIS(),
    )
    print(pipe,"\n")

    rank = pipe.evaluate(dm)
    print(rank,"\n")

    rankings = list(rank.values)
    print(rankings,"\n")

    for i in range(len(rankings)):
        sheet.cell(row= i+4, column=column).value = rankings[i]
    sheet.cell(row= 1, column=column).value = 'Rankings'
    wb.save(file_name)
