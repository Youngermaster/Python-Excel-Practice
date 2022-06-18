import openpyxl as xl
import os
from copy import copy

if __name__ == "__main__":
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    DOCUMENT_PATH = os.path.join(ROOT_DIR, 'documents')

    # opening the source excel file
    filename = os.path.join(DOCUMENT_PATH, 'trading.xlsx')
    wb1 = xl.load_workbook(filename)
    ws1 = wb1.worksheets[0]

    # opening the destination excel file
    filename1 = os.path.join(DOCUMENT_PATH, 'test.xlsx')
    wb2 = xl.load_workbook(filename1)
    ws2 = wb2.active

    # calculate total number of rows and
    # columns in source excel file
    mr = ws1.max_row
    mc = ws1.max_column

    # copying the cell values from source
    # excel file to destination excel file
    for i in range(1, mr + 1):
        for j in range(1, mc + 1):
            # reading cell value from source excel file
            c = ws1.cell(row=i, column=j)

            # writing the read value to destination excel file
            ws2.cell(row=i, column=j).value = c.value

    # saving the destination excel file
    wb2.save(str(filename1))
