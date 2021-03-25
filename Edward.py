from editpyxl import Workbook
'''
https://editpyxl.readthedocs.io/en/latest/usage.html
'''
wb = Workbook()

source_filename = r'A.xlsm'

wb.open(source_filename)

ws = wb["PARAMETER"]

ws.cell('A1').value = 3.14

wb.save('A2.xlsm')

wb.close()