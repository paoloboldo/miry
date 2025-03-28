
import openpyxl
import pandas as pd

file = r"assets\TURNI GIUGNO.xlsx"

wb = openpyxl.load_workbook(file)
sheet_names = wb.sheetnames
weeks = []
for name in sheet_names:
    if "sett" in name:
        weeks.append(name)
weeks.sort()

def complete_joint_cells(array):
    for i, value in enumerate(array):
        if "Unnamed" in value:
            array[i] = array[i-1]
    return array

# sheet_name = "Sett
df = pd.read_excel(file, sheet_name="Foglio3")
month_day = complete_joint_cells(df.columns.to_numpy()[1:])  # giorno del mese
values = df.values.to_numpy()
week_day = complete_joint_cells(values[0, 1:])  # giorno della settimana
worker = values[1, 1:]  # nome dipendente
times = values[:,0]
first_t = None
last_t = None
for t, time in enumerate(times):
    if "Unnamed" not in time and start_time is None:
        start_time = t
        continue
    if "Unnamed" not in time and start_time is None:




