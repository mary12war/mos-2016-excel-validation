from openpyxl import Workbook
from openpyxl import load_workbook

filename ="PathRockCrawlingP6-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Qtr 1"]
worksheet2 = wb["Qtr 2"]
ejemplo = open("TableList.txt", "a")

puntaje = 0

#Task1: Enable the Totals Row located in the “Qtr1” worksheet
tbl = worksheet1.tables["Tabla1"]
if (type(tbl).__name__ == "Table") and (worksheet1["E15"].value == "=SUBTOTAL(109,Tabla1[Total])"):
    puntaje +=1
else:
    print("No cumplió con TASK 1")

"""Task2: In the “Qtr1” worksheet, in the “Max” row, insert a formula in column B that
calculates the highest number of successful attempts for the month of January."""
maxfunct = worksheet1["B17"].value
if worksheet1["B17"].value == "=MAX(Tabla1[Jan])":
    puntaje +=1
else:
    print("No cumplió con TASK 2")

"""Task3: In the “Qtr1” worksheet, use the Path, Jan, Feb, and Mar columns to create a
Grouped 3D Bar chart. Do not include the “Total” column. Position the new chart to the
right of the table"""

"""Task4: In the “Qtr2” worksheet, create a table with cell range A9:E14 by applying the
Table Style Medium 18. Use the data in row 9 as headers"""
"""tbl2 = worksheet2.tables["Table2"]
tbls2 = str(tbl2).split(",")
#print(type(tbl2).__name__)
print(tbl2)
try:
    for i in tbls2:
        print("OK")
        if(i=="name='Table2"):
            print("OK2")
            if(i=="ref='A9:E14'"):
                print("OK3")
                if(i=="name='TableStyleMedium18'"):
                    print("OK4")
                    puntaje +=1
except KeyError:
    print("No cumplió con TASK 2")
    #print(tbl2[i],"WE WE WE\n\n\n")"""
#Task 5: Apply Style 3 to the pie chart in the “Qtr2” worksheet.

print("\nPuntaje de Proyecto A:", (puntaje),"/5")

#wb.save(filename)