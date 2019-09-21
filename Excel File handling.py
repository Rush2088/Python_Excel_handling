from cell_lib import list2col, list2row,read_range
from openpyxl import load_workbook
workbook_name = 'test.xlsx'
wb = load_workbook('test.xlsx')
ws = wb['Sheet1']

# New data to write:
data_list = [['100 Spit','38 Awaba','3 Anzac','13 Shadforth']]

####### Write function code snippets ################
# Write a column of values
# startcell= 'D14'   #enter starting cell number is A12 format
# list2col(startcell, data_list,ws)


# # Write a row of values
# startcell= 'B3'   #enter starting cell number is A12 format
# list2row(startcell, data_list,ws)

# wb.save(filename=workbook_name)

####### Read function code snippets ################
A=read_range('A1','B6',ws)
A_T=list(map(list, zip(*A))) # to transpose if required

