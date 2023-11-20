from openpyxl import Workbook
from openpyxl import load_workbook

filename ="EmployeeBonusesP7-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Employee Bonus"]
worksheet2 = wb["Sellers"]

puntaje = 0

"""Task1: Use Autofill to copy the formula from cell H4 that calculates “Total
Compensation” for each employee in the “Employee Bonuses” table"""
inicio = worksheet1["H5"].value
final = worksheet1["H11"].value
if inicio == "=SUM(F5:G5)" and final == "=SUM(F11:G11)":
    puntaje +=1
else:
    print("No cumplió con TASK 1")

"""Task2: Insert a formula into cell G4 that evaluates whether the number of “Parts, Accesory or Services” 
exceeds the Quarterly Goal. For each column that exceeds the goal, apply the Quarterly Bonus Rate."""
qGoal1 = worksheet1["G4"].value
qGoal2 = worksheet1["G11"].value
equaGoal = "=IF(Tabla1[[#This Row],[Total Sales]]>$B$18,Tabla1[[#This Row],[Total Sales]]*$B$17,0)"
if (qGoal1 and qGoal2) == equaGoal:
    puntaje +=1
else:
    print("No cumplió con TASK 2")    

#Task3: In the “Sellers” worksheet, delete the row that contains the “Allen” seller.
Allen = worksheet2["A11"].value
if Allen == "Allen":
    print("No cumplió con TASK 3")
else:
    puntaje +=1
#Task4: In the “Employee Bonuses” worksheet, disable the headers in the “Rates” table.
Head1 = worksheet1["A15"].value
Head2 = worksheet1["B15"].value
if Head1 == "Column1" and Head2 == "Column2":
    print("No cumplió con TASK 4")
else:
    puntaje +=1    
#Task 5: In cell F4 of the “Sellers” worksheet, insert an Jan to Mar line sparkline
print("Check for Sparkline line in worksheet")

print("\nPuntaje de Proyecto A:", (puntaje),"/5")

#wb.save(filename)