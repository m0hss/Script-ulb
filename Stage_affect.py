import openpyxl
from openpyxl.styles import Font
import time


def read_file(path):
    return openpyxl.load_workbook(path).active


def get_name(sheet, val):
    for i in range(2, sheet.max_row + 1):
        if sheet.cell(i, 1).value == val:
            return sheet.cell(i, 2).value
    return -1


def find_place_by(sheet, val1, val2):
    for i in range(2, sheet.max_row):
        if sheet_place.cell(i, 4).value != 0 and (val1 == sheet.cell(i, 2).value) and (val2 == sheet.cell(i, 3).value):
            sheet_place.cell(i, 4).value -= 1
            return True


def find_place(sheet):
    for j in range(2, sheet.max_row):
        if sheet_place.cell(j, 4).value != 0:
            sheet_place.cell(j, 4).value -= 1
            return j


def get_index_blank(sheet):
    for i in range(2, sheet_pref.max_row):
        if not sheet_pref.cell(i, sheet_pref.max_column).value:
            return i


sheet_pref = read_file("Annexe 1 - préférences.xlsx")
sheet_rang = read_file("Annexe 2 - classement.xlsx")
sheet_hopital = read_file("Annexe 3 - hopitaux.xlsx")
sheet_service = read_file("Annexe 4 - services.xlsx")
sheet_place = read_file("Annexe 5 - places.xlsx")
sheet_pref.delete_rows(4223, sheet_pref.max_row)
stage = {}
matricules = []
start_time = time.time()

for row in range(2,  237+ 1):
    matricule = (sheet_rang.cell(row, 1).value)
    
    
    
    hopital = []
    service = []
    typePref = []
    pref = {}
    flag = False
    c = 0

    for row1 in range(2, sheet_pref.max_row):
        if matricule == sheet_pref.cell(row1, 3).value:
            pref[sheet_pref.cell(row1, 6).value] = sheet_pref.cell(row1, 1).value

    pref = dict(sorted(pref.items()))

    for pref_id in pref.values():
        for row2 in range(2, sheet_pref.max_row):
            if pref_id == sheet_pref.cell(row2, 1).value:
                hopital.append(sheet_pref.cell(row2, 4).value)
                service.append(sheet_pref.cell(row2, 5).value)
                typePref.append(sheet_pref.cell(row2, 7).value)

    for h, s, t in zip(hopital, service, typePref):
        c += 1
        if t == 1 or c == len(pref):
            if find_place_by(sheet_place, h, s):
                stage.setdefault("Matricule", []).append(matricule)
                stage.setdefault("Hopital", []).append(get_name(sheet_hopital, h))
                stage.setdefault("Service", []).append(get_name(sheet_service, s))
                flag = True
                break
    if not pref or flag or matricule not in stage['Matricule']:
        if find_place(sheet_place):
            stage.setdefault("Matricule", []).append(matricule)
            stage.setdefault("Hopital", []).append(get_name(sheet_hopital, sheet_place.cell(find_place(sheet_place), 2).value))
            stage.setdefault("Service", []).append(get_name(sheet_service, sheet_place.cell(find_place(sheet_place), 3).value))


    print('---------------------------------------------------------------')

print(stage)

wb_stage = openpyxl.Workbook()
sheet_stage = wb_stage.active
sheet_stage["A1"] = "Matricule"
sheet_stage["A1"].font = Font(size=12, bold=True)
sheet_stage.column_dimensions['A'].width = 10
sheet_stage["B1"] = "Hopital"
sheet_stage["B1"].font = Font(size=12, bold=True)
sheet_stage.column_dimensions['B'].width = 30
sheet_stage["C1"] = "Service"
sheet_stage["C1"].font = Font(size=12, bold=True)
sheet_stage.column_dimensions['C'].width = 20

row = 1
for val in stage.values():
    for j in range(0, len(val)):
        sheet_stage.cell(j + 2, row).value = val[j]
    row += 1
print("--- %s seconds ---" % (time.time() - start_time))
wb_stage.save("Stages_Affectations.xlsx")