import zlrd
wb = xlrd.open_workbook(<filename>, formatting_info=True)
ws = wb.sheet_by_name(<sheet name>)
c = ws.cell(1, 1)
cif = ws.cell_xf_index(1, 1)
iif = wb.xf_list[cif]
cbg = iif.background.pattern_colour_index
print(cbg)