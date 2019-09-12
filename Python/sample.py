#from excel2json import convert_from_file
import excel2json
import xlrd
from collections import OrderedDict
import json

EXCEL_FILE = './death-data.xlsx'  # or '../example.xlsx'

 
wb = xlrd.open_workbook(EXCEL_FILE)


print(wb.sheet_names())


sh = wb.sheet_by_index(0)
# print(sh.cell(4,5).color_index)
# xfx = sh.cell_xf_index(4,5)
# xf = wb.xf_list(xfx)
# bgx = xf.background.pattern_colour_index
# print(bgx)
# pattern_colour = wb.colour_map(bgx)
# print(pattern_colour)



#EXCEL to JSON 


wb = xlrd.open_workbook(EXCEL_FILE)
sh = wb.sheet_by_index(2)
rowval = sh.row_values(7)
print(rowval[3])

main=OrderedDict()
for x in range(1,sh.nrows):
   rowval = sh.row_values(x)
   if(rowval[0]!=''):
        sub=OrderedDict()
        for y in range(1,sh.ncols):
            if(rowval[y]!=''):
                for z in range(y,sh.ncols)
                    rowval[y+1]

#     if(rowval[y]==''):
#      print(x, y, "hello")
#     else:
#         print(x, y, "bye")
# if(sh.row_values(rownum)1)

print(sh.row_values(0))
print(sh.nrows)
# List to hold dictionaries
cars_list = []
# Iterate through each row in worksheet and fetch values into dict
# for rownum in range(1, sh.nrows):
#     cars = OrderedDict()
#     row_values = sh.row_values(rownum)
#     cars['car-id'] = row_values[0]
#     cars['make'] = row_values[1]
#     cars['model'] = row_values[2]
#     cars['miles'] = row_values[3]
#     cars_list.append(cars)
    



# print(cars)
    # print(cars_list)
# Serialize the list of dicts to JSON
j = json.dumps(cars_list)
# Write to file
with open('data.json', 'w') as f:
    f.write(j)




print("Hello")