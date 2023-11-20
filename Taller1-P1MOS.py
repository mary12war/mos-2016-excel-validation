from openpyxl import Workbook
from openpyxl import load_workbook

filename ="RentalAccessories-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Sheet1"]
worksheet2 = wb["Boat"]
worksheet3 = wb["Inv"]

puntaje = 0

#Task1: Starting with the cell A16 in Sheet1, import the list from the delimited by commas  file “More.csv” (Accept the other predetermined options).
inicio = worksheet1["A16"].value
final = worksheet1["F23"].value
if inicio == 11 and final == 750:
    puntaje +=1
else:
    print("No cumplió con TASK 1")

#Task2: Sheet1 Color HEX 2F75B5 - Blue, Accent 1, darkness 25%
tabCol = worksheet1.sheet_properties.tabColor.tint
if tabCol == -0.249977111117893:
    puntaje +=1
else:
    print("No cumplió con TASK 2")

#Task3: Copy the content of the worksheet “Inv” and place it in the table of the “Boat” worksheet starting at cell A6
inicio = worksheet2["A6"].value
final = worksheet2["D13"].value
if inicio == 10 and final == 1750:
    puntaje +=1
else:
    print("No cumplió con TASK 3")

#Task4:  Rearrange the worksheets so “Boat” is the first one.
position = wb._sheets[0]
print(str(position))
if position == "Boat":
    puntaje +=1
else:
    print("No cumplió con TASK 4")

#Task 5. Hide the “Inv” worksheet.
if worksheet3.sheet_state == "hidden":
    puntaje +=1
else:
    print("No cumplió con TASK 5")

print("\nPuntaje de Proyecto A:", (puntaje),"/5")

wb.save(filename)