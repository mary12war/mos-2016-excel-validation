from openpyxl import Workbook
from openpyxl import load_workbook
import os

dir = os.getcwd()
#print(dir)
ext =(".xlsx")
punt = open("puntajesAM.txt", "a")
count = 0
punt.write(f"Estudiante #; Nombre; Q1; Q2; Q3; Q4; Q5; Q6; Q7; Q8; Q9; Total; Porcentaje\n")
for files in os.listdir(dir):
    if files.endswith(ext):
        filename = files
        wb = load_workbook(filename)
        sheet1 = wb["INSTRUCTION"]
        sheet2 = wb["Faculty Advisors"]
        sheet3 = wb["Faculty List"]
        sheet4 = wb["Part-Time"]
        sheet5 = wb._sheets[4]
        sheet5_name = str(sheet5)
        name = sheet1["B3"].value
        puntaje = 0

        #PREGUNTA 1 
        cellphone = sheet2["L1"].value
        if cellphone == "Cell Phone":
            puntaje +=1
            ok1 = 1
        else:
            ok1 = 0
            punt.write("Pregunta 1 No Cumplida\n")
            print("Pregunta 1 No Cumplida")

        #PREGUNTA 2
        viejo1 = sheet3["A20"].value
        viejo2 = sheet3["C30"].value
        nuevo1 = sheet4["A2"].value
        nuevo2 = sheet4["C12"].value
        if (viejo1 and viejo2) == None:
            if nuevo1 == "Sideris" and nuevo2 == "Part-Time":
                    puntaje +=1
                    ok2 = 1
            else:
                ok2 = 0
                punt.write("Pregunta 2 No Cumplida\n")
                print("Pregunta 2 No Cumplida")
        else:
            ok2= 0
            punt.write("Pregunta 2 No Cumplida\n")
            print("Pregunta 2 No Cumplida")
        ok3 = "TBD"
        ok4 = "TBD"
        #PREGUNTA 5
        if sheet5_name == '<Worksheet "Menu Sales">':
            ok5= 1
            puntaje +=1
        else:
            ok5= 0
            punt.write("Pregunta 5 No Cumplida\n")
            print("Pregunta 5 No Cumplida")
        ok6= "TBD"
        ok7= "TBD"
        #PREGUNTA 8
        aver = sheet2["Q2"].value
        if aver == "=AVERAGE(F2:F30)" or aver =="=AVERAGE(F1:F30)":
            ok8= 1
            puntaje +=1
        else:
            ok8= 0
            punt.write("Pregunta 8 No Cumplida\n")
            print("Pregunta 8 No Cumplida")

        #PREGUNTA 9
        cnt = sheet2["Q3"].value
        if cnt == "=COUNT(K1:K30)" or cnt == "=COUNT(K2:K30)":
            ok9= 1
            puntaje +=1
        else:
            ok9= 0
            punt.write("Pregunta 9 No Cumplida\n")
            print("Pregunta 9 No Cumplida")

        print("Estudiante:", name, "\nPuntaje:", (puntaje/9*100))
        total = (puntaje/9)*10
        punt.write(f"Estudiante {count}; {name}; {ok1}; {ok2}; {ok3}; {ok4}; {ok5}; {ok6}; {ok7}; {ok8}; {ok9}; {puntaje}; {total}\n")
    punt.close()