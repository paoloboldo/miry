import json
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import datetime

def get_week_days(num_settimana, anno):
    import datetime
    # Definisco l'ordine dei giorni della settimana
    giorni = ["LUN", "MAR", "MER", "GIO", "VEN", "SAB", "DOM"]
    
    # Ottengo il lunedì della settimana ISO (anche se la data potrebbe appartenere all'anno precedente)
    lunedi = datetime.date.fromisocalendar(anno, num_settimana, 1)
    
    # Creo il dizionario: per ogni giorno della settimana, prendo il giorno del mese (day)
    giorni_mese = {giorni[i]: (lunedi + datetime.timedelta(days=i)).day for i in range(7)}
    
    return giorni_mese


def crea_settimanale(num_settimana=23, anno=2025):

    # Carica i dati dal file impostazioni.json
    with open("impostazioni.json", "r") as f:
        settings = json.load(f)
    dipendenti = settings["dipendenti"]
    
    # Giorni da visualizzare nelle intestazioni (in ordine)
    giorni = ["LUN", "MAR", "MER", "GIO", "VEN", "SAB", "DOM"]

    # Prendo il giorno del mese per ogni giorno della settimana
    giorni_mese = get_week_days(num_settimana, anno)
    
    # Crea una nuova cartella Excel e seleziona il foglio attivo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sett1"
    
    # Imposta un bordo sottile da applicare alle celle
    thin_border = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )
    
    # 1) Colonna A: etichetta "ORARIO" su righe 1-3
    ws.merge_cells("A1:A3")
    cell_orario = ws["A1"]
    cell_orario.value = "ORARIO"
    cell_orario.alignment = Alignment(
        horizontal="center", vertical="center"
    )
    # Ad esempio sfondo rosso e testo bianco, per imitare l'immagine
    cell_orario.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    cell_orario.font = Font(color="FFFFFF", bold=True)
    
    # Imposta larghezza colonna A
    ws.column_dimensions["A"].width = 8
    
    # 2) Inserisci gli orari (7:30 -> 20:30 ogni 30 min) in colonna A dalla riga 4 in poi
    start_time = datetime.datetime(2023, 1, 1, 7, 30)
    end_time = datetime.datetime(2023, 1, 1, 20, 30)
    delta = datetime.timedelta(minutes=30)
    
    current = start_time
    row_idx = 4
    while current <= end_time:
        ws.cell(row=row_idx, column=1).value = current.strftime("%H:%M")
        row_idx += 1
        current += delta
    
    # 3) Struttura delle colonne per i giorni
    #    Ogni giorno occupa tante colonne quanti sono i dipendenti in impostazioni.json "dipendenti"[]
    #    Giorno 1 -> colonne B, C, D  (start_col=2, end_col=4)
    #    Giorno 2 -> colonne E, F, G  (start_col=5, end_col=7)
    #    ...
    #    In generale: day i (1-based) -> start_col = 2 + num_dipendenti*(i-1), end_col = start_col + num_dipendenti - 1
    num_dipendenti = len(dipendenti)

    
    for i in range(7): # includiamo la domenica (0..6)
        # i = 0 -> LUN, i = 1 -> MAR, ..., i = 6 -> DOM
        day_name = giorni[i]   # LUN, MAR, ...
        day_num = giorni_mese[day_name]
        
        start_col = 2 + num_dipendenti*i
        end_col = start_col + num_dipendenti - 1  # Adjusted to use num_dipendenti
        
        # Riga 1: celle unite con il numero del giorno
        ws.merge_cells(
            start_row=1, start_column=start_col,
            end_row=1, end_column=end_col
        )
        cell_day_num = ws.cell(row=1, column=start_col)
        cell_day_num.value = str(day_num)
        cell_day_num.alignment = Alignment(horizontal="center", vertical="center")
        # Testo rosso in grassetto (come da screenshot)
        cell_day_num.font = Font(color="FF0000", bold=True)
        
        # Riga 2: celle unite con il nome del giorno
        ws.merge_cells(
            start_row=2, start_column=start_col,
            end_row=2, end_column=end_col
        )
        cell_day_name = ws.cell(row=2, column=start_col)
        cell_day_name.value = day_name
        cell_day_name.alignment = Alignment(horizontal="center", vertical="center")
        cell_day_name.font = Font(bold=True)
        
        # Colore di sfondo azzurrino per i giorni pari
        if day_num % 2 == 0:
            fill_color = PatternFill(start_color="CCFFFF", end_color="CCFFFF", fill_type="solid")
            # Imposta il colore del testo a nero
            day_num_color = "000000"
            day_name_color = "000000"
        else:
            fill_color = None
            # Imposta il colore del testo a rosso
            day_num_color = "FF0000"
            day_name_color = "FF0000"

        # Applica il colore del testo e il grassetto al numero del giorno
        cell_day_num.font = Font(color=day_num_color, bold=True)

        # Applica il colore del testo e il grassetto al nome del giorno
        cell_day_name.font = Font(color=day_name_color, bold=True)
        
        # Applica il bordo e l’eventuale sfondo a tutte le righe interessate (1..30) in queste colonne
        for r in range(1, 31):
            for c in range(start_col, end_col + 1):  # Adjusted to use num_dipendenti
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                # Se vuoi colorare anche intestazioni (riga 1 e 2), togli la condizione (r >= 4)
                if fill_color and r >= 4:
                    cell.fill = fill_color
        
        # Imposta larghezza delle colonne
        for c in range(start_col, end_col + 1):  # Adjusted to use num_dipendenti
            ws.column_dimensions[get_column_letter(c)].width = 5
    
    # 4) Inserisci i nomi dei dipendenti, uno per colonna per ogni giorno
    for i in range(7):  # Per ogni giorno della settimana
        day_name = giorni[i]
        start_col = 2 + num_dipendenti * i  # Adjusted to use num_dipendenti

        for j, dip in enumerate(dipendenti):  # Per ogni dipendente
            col = start_col + j
            cell_emp = ws.cell(row=3, column=col)  # Modificato a riga 3
            cell_emp.value = dip["nome"]
            cell_emp.alignment = Alignment(horizontal="center", vertical="center")

            # Applica il colore di sfondo definito nel JSON (campo "colore")
            color_code = dip.get("colore", None)
            if color_code:
                fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")
                cell_emp.fill = fill
    
    # 5) Applica i bordi anche alla colonna A (orari)
    for r in range(1, 31):
        cell = ws.cell(row=r, column=1)
        cell.border = thin_border
    
    # (Facoltativo) Se vuoi un'ulteriore riga di "chiusura" (riga 31) con bordi
    max_col = 1 + num_dipendenti * 7  # Adjusted to use num_dipendenti
    for c in range(1, max_col + 1):
        cell = ws.cell(row=31, column=c)
        cell.border = thin_border
    
    # Salva il file risultante
    wb.save("Sett1.xlsx")

if __name__ == "__main__":
    crea_settimanale()
