import openpyxl, json, re, glob, sys
from datetime import datetime

SHEET_NAMES = ['Genel','Genel Müşteri Çıkış','Adresleme','Toplama','Sorter','Tekli-Eksik Sevk','Kargo','Adres Denetim','İade Kabul']
AY_ISIMLERI = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']

xlsx_files = glob.glob('*.xlsx')
if not xlsx_files:
    print("HATA: Repo içinde .xlsx dosyası bulunamadı.")
    sys.exit(1)
xlsx_path = xlsx_files[0]
print(f"Excel dosyası bulundu: {xlsx_path}")

wb = openpyxl.load_workbook(xlsx_path, data_only=True)

def to_ymd(d):
    return f"{d.year:04d}-{d.month:02d}-{d.day:02d}"

newData = {}

for sName in SHEET_NAMES:
    if sName not in wb.sheetnames:
        continue
    ws = wb[sName]
    rows = list(ws.iter_rows(values_only=True))
    n = len(rows)

    blocks = []
    for i in range(n):
        row = rows[i]
        if row and len(row) > 1 and row[1] == 'Tarih':
            dates = [v for v in row[2:] if isinstance(v, datetime)]
            if dates:
                year_counts = {}
                for d in dates:
                    year_counts[d.year] = year_counts.get(d.year, 0) + 1
                majority_year = max(year_counts, key=year_counts.get)
                blocks.append({'dateRow': i, 'haftaRow': i+1, 'saatRow': i+2, 'adetRow': i+3, 'year': majority_year})

    newData[sName] = {}
    for b in blocks:
        yr = b['year']
        sfx = str(yr)[-2:]
        date_row = rows[b['dateRow']]
        hafta_row =
      
