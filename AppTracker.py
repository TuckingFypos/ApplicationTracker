import xlwt
import xlrd

workbook = xlrd.open_workbook('ApplicationDatabase.xlsx')
sheet = workbook.sheet_by_index(0)

keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

dict_info = []
for row_index in range(1, sheet.nrows):
    d = {keys[col_index]: sheet.cell(row_index, col_index).value
         for col_index in range(sheet.ncols)}
    dict_info.append(d)

print(dict_info)
