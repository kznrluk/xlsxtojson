##
# xlsxtojson.py Convert .xlsx file to JSON.
##

import xlrd
import sys
import re
import json

if len(sys.argv) is 1:
    print('ファイルの指定がありません。')
    print('Use : xlsxtojson.py [.xlsx file]')
    exit()

Book = xlrd.open_workbook(sys.argv[1])
sheet_1 = Book.sheet_by_index(0)
json_data = {}
keys = []
rows = []
cols = []

for row in range(sheet_1.nrows):
    for col in range(sheet_1.ncols):
        # 列情報
        if row is 0:
            if col is not 0:
                rows.append(sheet_1.cell(row, col).value)
        # 行情報
        else:
            if col is 0:
                cols.append([])
                cols[row-1].append(re.sub(r'[^A-z]+','',sheet_1.cell(row, col).value))
            else:
                cols[row-1].append(sheet_1.cell(row, col).value)

for i,col in enumerate(cols):
    for k,value in enumerate(col):
        if k is 0:
            json_data[value] = {}
            key = value
        else:
            json_data[key][rows[k-1]] = value


if len(sys.argv) is 3:
    jsonfile = open(sys.argv[2],'w')
    json.dump(json_data,jsonfile,indent=4)
else:
    print("{}".format(json.dumps(json_data,indent=4)))