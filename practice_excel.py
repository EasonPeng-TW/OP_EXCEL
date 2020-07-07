import openpyxl

workbook = openpyxl.load_workbook('test1.xlsx')
sheet = workbook.worksheets[0]
cell = sheet['A1']
	# sheet['A2'] ='456'
print(cell.value)
if cell.value == 123:
    sheet['A2'] ='456'
workbook.save("test1.xlsx")