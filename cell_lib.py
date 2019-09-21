import re
# from openpyxl import load_workbook


def cellconv(cell):
    row_number = int(re.sub('[^0-9]+', '', cell))
    col_ = re.sub('[^a-zA-Z]+', '', cell)
    col_number = alpha2num(col_)
#   col_number =ord(col_.lower())- 96
    return col_number,row_number

def alpha2num(alphaN):
    '''
    base conversion - 26 to decimal
    '''
    num=0
    for i in range (len(alphaN)):
        x=-(i+1)
        y=26 ** (-x-1)
        num+= y* int(ord(alphaN[x].lower())- 96)
    return num



def list2col(startcell, lst,ws):
    '''
    Write a list to column in excel
    '''
    start_col ,start_row= cellconv(startcell)   # starting cell number is A12 format
    
    for column, column_entries in enumerate(lst, start=start_col):
        for row, value in enumerate(column_entries, start=start_row):
            ws.cell(column=column, row=row, value=value)



def list2row(startcell, lst,ws):    
    '''
    Write a row of values in excel
    '''
    start_col ,start_row= cellconv(startcell)   # starting cell number is A12 format

    for row, row_entries in enumerate(lst, start=start_row):
        for column, value in enumerate(row_entries, start=start_col):
            ws.cell(column=column, row=row, value=value)



def read_range(x,y,ws):
    cells = ws[x:y]  # Cell range to read from excel
    rows , cols = len(cells), len(cells[0])
    A = [ [ None for i in range(cols) ] for j in range(rows) ]

    for r in range(rows):
        for c in range(cols):
            A[r][c]=cells[r][c].value
    return A