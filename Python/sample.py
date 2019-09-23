#from excel2json import convert_from_file
import excel2json
import xlrd
from collections import OrderedDict
import json
import pymongo

EXCEL_FILE = './death-data.xlsx'  # or '../example.xlsx'

 
wb = xlrd.open_workbook(EXCEL_FILE)


print(wb.sheet_names())


sh = wb.sheet_by_index(0)

level=1
for rownum in range(13,sh.nrows):
    rowval = sh.row_values(rownum)
    nextrowval = sh.row_values(rownum+1)
    for colnum in range(1,5):
        if(rowval[colnum]!=''):
            if(nextrowval[colnum]!=''):
                level=level+1
                break

print(level)
# print(sh.cell(4,5).color_index)1
# xfx = sh.cell_xf_index(4,5)
# xf = wb.xf_list(xfx)
# bgx = xf.background.pattern_colour_index
# print(bgx)
# pattern_colour = wb.colour_map(bgx)
# print(pattern_colour)



#EXCEL to JSON 


# wb = xlrd.open_workbook(EXCEL_FILE)
# sh = wb.sheet_by_index(2)
# countries=sh.row_values(0)
# rowval = sh.row_values(7)
# print("hello")
# print(rowval[2])
# rowval[2]=OrderedDict()
# Communicable, maternal, perinatal and nutritional conditions.append({"name":"shreya"})
# print(rowval[2])


# main=OrderedDict()
# level=1
# for x in range(1,sh.nrows):
#    rowval = sh.row_values(x)
#    if(rowval[0]!=''):
#         # sub=OrderedDict()
#         for y in range(1,5):

#             if(rowval[y]!=''):
#                 # break after values appended
#                 main.append(7)
#                 for z in range(y,sh.ncols):

#                     # 
#                     obj=OrderedDict()
#                     obj['country']=countries[y]
#                     obj['deaths']=rowval[y]
#                     main.append(obj)
                
#                 else:
#                     level=level+1
#                 for z in range(y,sh.ncols)
#                    sub[]=rowval[y]

# #     if(rowval[y]==''):
# #      print(x, y, "hello")
# #     else:
# #         print(x, y, "bye")
# # if(sh.row_values(rownum)1)

# print(sh.row_values(0))
# print(sh.nrows)
# # List to hold dictionaries
# cars_list = []
# # Iterate through each row in worksheet and fetch values into dict
# # for rownum in range(1, sh.nrows):
# #     cars = OrderedDict()
# #     row_values = sh.row_values(rownum)
# #     cars['car-id'] = row_values[0]
# #     cars['make'] = row_values[1]
# #     cars['model'] = row_values[2]
# #     cars['miles'] = row_values[3]
# #     cars_list.append(cars)
    



# # print(cars)
#     # print(cars_list)
# # Serialize the list of dicts to JSON
# j = json.dumps(cars_list)
# # Write to file
# with open('data.json', 'w') as f:
#     f.write(j)




# print("Hello")



# send data to mongo db--------------------------------- 

# client = MongoClient('localhost', 27017)
# db = client['countries_db']
# collection_currency = db['currency']

# with open('currencies.json') as f:
#     file_data = json.load(f)

# # use collection_currency.insert(file_data) if pymongo version < 3.0
# collection_currency.insert_one(file_data)  
# client.close()