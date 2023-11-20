from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.excel import Guid

filename ="BicycleSaleP5-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Sales Q1"]
worksheet2 = wb["List"]

puntaje = 0

#Task1:  Set up the “Sales Q1” worksheet so that rows 1 through 3 remain visible as you scroll vertically.
frzpanes = worksheet1.freeze_panes
if frzpanes != None:
    puntaje +=1
else:
    print("No cumplió con TASK 1")

#Task2: Add the “Team Proposals” subject to the document properties.
#propert = wb.appName 
#print(propert)

#Task3: Configure Excel to always print the range of A1:F17 cells in the worksheet“Sales Q1.”

#Task4: In the “Sales Q1” worksheet, insert a function in cell B19 that calculates all sales in the “Total” column.
suma = "=SUM(Tabla2[Total])"
if worksheet1["B19"].value == suma:
    puntaje +=1
else:
    print("No cumplió con TASK 4")
    
"""Task 5. In cell B4 of worksheet “Sales Q1”, insert a function that joins the “Description 
and Style” catalog separated by a dash. Include a space on both sides of the script 
(Example: “Cross country – Stiff”)."""
answer = '=CONCATENATE(Tabla3[[#This Row],[Description]]," - ",Tabla3[[#This Row],[Style]])'
if (worksheet1["B4"].value == answer) and (worksheet1["B17"].value == answer):
    puntaje += 1
else:
    print("No cumplió con TASK 5")

print("\nPuntaje de Proyecto A:", (puntaje),"/5")

#wb.save(filename)