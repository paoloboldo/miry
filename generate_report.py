
import openpyxl
import pandas as pd

file = r"assets\TURNI GIUGNO.xlsx"

# workbook = openpyxl.Workbook()
# sheet = workbook.active

df = pd.read_excel(file, sheet_name="Foglio3")
col_1 = df.columns
col_2 = df.values[0]
col_3 = df.values[1]
orari = df.values[:,0]

