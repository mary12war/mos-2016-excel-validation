from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl_image_loader import SheetImageLoader

filename ="P4-Solution.xlsx"

wb = load_workbook(filename)
sheet = wb.active
worksheet1 = wb["Hardware"]


puntaje = 0

#Task1: Add a new worksheet called “Customers” to the workbook.
try:
    new_ws1 = wb._sheets[0]
    new_ws2 = wb._sheets[1]
    if new_ws1 or new_ws2 == "Customers":
        puntaje +=1
    else:
        print("No cumplió con TASK 1")
except IndexError:
    print("No cumplió con TASK 1")

#Task2: Simultaneously remove all duplicate records in the “Wired networks” table
dupl = 2356
if ((worksheet1["A6"].value and worksheet1["A7"].value) == dupl):
    print("No cumplió con TASK 2")
elif ((worksheet1["A10"].value and worksheet1["A14"].value) == dupl):
    print("No cumplió con TASK 2")
else:
    puntaje +=1
#Task3: Starting in cell A1 in the “Hardware” worksheet, import the image “NetworkTopology.png”
try:
    image_loader = SheetImageLoader(worksheet1)
    imagen = image_loader.get("A1")
    if type(imagen.show).__name__ == "method":
        puntaje +=1
except ValueError:
    print("No cumplió con TASK 3")
#Task4: In the Hardware worksheet, rotate the text “Wired networks” and “Wireless networks” with a Descending Angle

#Task 5.  Sort the data in the “Wired networks” table. Sort by “IDProduct”, from lowest to highest.
if ((worksheet1["A5"].value == 2356)and (worksheet1["A9"].value == 5847) and (worksheet1["A13"].value == 12343)):
    puntaje +=1
else:
    print("No cumplió con TASK 5")

print(f"\nPuntaje de Proyecto A: {puntaje}/5")

#wb.save(filename)