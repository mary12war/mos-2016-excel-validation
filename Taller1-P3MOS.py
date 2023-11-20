from openpyxl import Workbook
from openpyxl import load_workbook

filename ="TALLER 1 - CHECK/Project 3-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Cost"]
worksheet2 = wb["Profits"]

puntaje = 0

#Task1:Simultaneously replace all concurrency of the word “Choco” with “Chocolate” in the workbook.
wch = "Waste of Chocolate from heaven"
chmwc = "Chocolate mint with cherry"
spcych = "Spicy Chocolate"
cmch = "Cherry Mint Chocolate"
chclt = "Chocolate"
if ((worksheet1["A11"].value and worksheet2["A15"].value == wch) 
    and (worksheet1["A19"].value ==chmwc) 
    and (worksheet1["A21"].value and worksheet2["A14"].value ==spcych) 
    and (worksheet2["A22"].value == cmch) 
    and (worksheet1["A29"].value == chclt)):
    puntaje +=1
else:
    print("No cumplió con TASK 1")

#Task2:
#Task3  In cell B28 of the “Profits” worksheet, insert a formula that shows the number of “Sales” greater than 250.:
countif_flavor = worksheet2["B28"].value
if countif_flavor == '=COUNTIF(Tabla2[Sales],">250")':
    puntaje += 1
else:
    print("No cumplió con TASK 3")    
#Task4:
#Task5: 

print("\nPuntaje de Proyecto A:", (puntaje),"/5")

#wb.save(filename)