import numpy as np
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
        if str(value) == "nan":
            array[i] = array[i-1]
    return array


def first_last_indices(array):
    first_i = None
    last_i = None
    for i, value in enumerate(array):
        if str(value) != "nan":
            last_i = i
            if first_i is None:
                first_i = i
            continue
        else:
            if first_i is not None:
                break
    return [first_i, last_i]


# sheet_name = "Sett
df = pd.read_excel(file, sheet_name="Foglio3")
values = df.values

worker = values[1]  # nome dipendente
first_w, last_w = first_last_indices(worker)
worker = worker[first_w: last_w + 1]

month_day = complete_joint_cells(df.columns.to_numpy()[first_w: last_w + 1])  # giorno del mese

week_day = complete_joint_cells(values[0, first_w: last_w + 1])  # giorno della settimana

times = values[:,0]
first_t, last_t = first_last_indices(times)
times = times[first_t: last_t + 1]

values = values[first_t: last_t + 1, first_w: last_w + 1]



