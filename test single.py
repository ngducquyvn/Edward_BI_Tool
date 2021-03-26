

#https://openpyxl.readthedocs.io/en/stable/tutorial.html
from editpyxl import Workbook

wb = Workbook()
'''

DESIGN = [
[1
,1]
, 
[2
,1]
]

for i in DESIGN:
	
	#print(i)

	#Full_Path_Report =  os.path.join(i[5], "A.XLSM")

	wb.open(r'B.xlsx')
	ws_par = wb["PARAMETER"]
	row_x = i[0]
	cel_x = i[1]

	print(cel_x)

	print (ws_par.cell(row=2,column = 1).value)
	print (ws_par.cell(row=row_x,column = cel_x).value)

	ws_par.cell(row=row_x,column = cel_x).value = "A"


	wb.save("D:\\Quynd4 Data\\! Build BI Tool\\Input\\B2.xlsx")
	#wb.close()
	
