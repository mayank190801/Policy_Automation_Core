#In this I am trying to put data into excel file as tables and everything
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
treeData = [["Type", "Leaf Color", ""], ["Maple", "Red", 549], ["Oak", "Green", 783], ["Pine", "Green", 1204]]

for row in treeData: 
    ws.append(row)


ft = Font(bold=True)
for row in ws["A1:C1"]:
    for cell in row:
        cell.font = ft


wb.save("Store.xlsx")