import os
import openpyxl
import skcriteria as skc
from skcriteria.preprocessing import invert_objectives, scalers
from skcriteria.madm import similarity
from skcriteria.pipeline import mkpipe
import datetime

from util import *



if __name__ == "__main__":
    file = get_file_path()
    wb = file[0]
    file_name = file[1]
    sheet = wb.active
    #getting ready to get info except matrix
    criteria = []
    weights = []
    objectives = []
    alternatives = []
    column = 2
    row = 4
    rank_column = 0
    while True:
        #getting all the values except matrix

        #criteria
        criteriaValue = sheet.cell(row= 1, column= column).value
        criteria.append(criteriaValue)

        #weights
        weightValue = sheet.cell(row= 2, column= column).value
        try:
            weights.append(float(weightValue))
        except:
            print('Incorrect weights',"\n")
            break
        
        # objectives (min, max)
        objectiveValue = sheet.cell(row= 3, column= column).value
        if objectiveValue == 'min':
            objectives.append(min)
        elif objectiveValue == 'max':
            objectives.append(max)

        #checking next cell
        column += 1
        #this will be used if the program has already been used on the file
        current_cell = sheet.cell(row= 1, column= column).value
        if current_cell != None:
            # seeing if next sell is a ranking
            if "Rankings by" in current_cell:
                rank_column = column
                while True:
                    rank_column += 1                
                    if not sheet.cell(row= 1, column= rank_column).value or not "Rankings by" in sheet.cell(row= 1, column= rank_column).value:
                        #getting compared object names (alternatives)
                        while True:
                            alternativeValue = sheet.cell(row= row, column= 1).value
                            alternatives.append(alternativeValue)
                            row += 1
                            if not sheet.cell(row= row, column= 1).value:
                                print(alternatives,"\n")
                                break 
                        break
                break
        #this will be used if the program is beeing used on the file for the first time
        else:
            if not current_cell or not sheet.cell(row= 2, column= column).value or not sheet.cell(row= 3, column= column).value:
                print(criteria, weights, objectives)
                #getting compared object names (alternatives)
                while True:
                    alternativeValue = sheet.cell(row= row, column= 1).value
                    alternatives.append(alternativeValue)
                    row += 1
                    if not sheet.cell(row= row, column= 1).value:
                        print(f'alternatives (compared objects): {alternatives}')
                        break
                rank_column = column
                break

    #getting matrix
    matrix = create_matrix(sheet, row, column)

    # evaluating values, putting results them in the file
    try:
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
        print(f"rank: {rank} \n")

        rankings = list(rank.values)
        print(f"rankings: {rankings} \n")

        #inserting ranks
        for i in range(len(rankings)):
            sheet.cell(row= i+4, column=rank_column).value = rankings[i]

        #inserting rank header with date & time
        sheet.cell(row= 1, column=rank_column).value = f'Rankings by {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}'

        #getting column name in letters
        index = column_index_to_letters(rank_column)

        #adjusting the width of the column with the results
        wb.worksheets[0].column_dimensions[index].width = 30
    except Exception as error:
        print("Processes stopped: Your file doesn't oblige by the structure given in README.md. Review for any possible holes or incorrect value types.")
        print(f"Error: {error}")

    #saving file
    wb.save(file_name)
    
