import openpyxl
import requests
import time

API_KEY = "6783637426dd4a501832b9e037b01631qCyf2L9b7a"
FILE_NAME = r"S:\Doc_Schad\LISTE\concordanza\collins_spider_list.xlsx" #file path 

# file loading
wb = openpyxl.load_workbook(FILE_NAME)
sheet = wb.active # Prende il primo foglio
 
print("Starting non-destructive check...")

# Partiamo dalla riga 2 (assumendo che la 1 sia l'header)
# max_row ci dice quante righe ci sono nel file
for row in range(3, sheet.max_row + 1):
    taxon = sheet.cell(row=row, column=4).value # Col D for species
    if not taxon:
        continue
        
    print(f"Checking row {row}: {taxon}")
    
    url = f"https://wsc.nmbe.ch/api/lsid?apiKey={API_KEY}&name={taxon}"
    try:
        r = requests.get(url, timeout=10).json()
        
        # Scriviamo direttamente nelle celle senza toccare il resto del file
        if r.get('valid') == 1:
            sheet.cell(row=row, column=8).value = str(taxon) # Colonna H
            sheet.cell(row=row, column=9).value = "OK"        # Colonna I
        elif r.get('currentValidName'):
            sheet.cell(row=row, column=8).value = r.get('currentValidName')
            sheet.cell(row=row, column=9).value = "OK"
        else:
            sheet.cell(row=row, column=8).value = "Not found"
            sheet.cell(row=row, column=9).value = "MISSING"
    except:
        sheet.cell(row=row, column=9).value = "ERROR"
    
    time.sleep(0.4)

# Salviamo il file: openpyxl manterrà la formattazione originale delle altre celle
wb.save(FILE_NAME)
print("Done!")

