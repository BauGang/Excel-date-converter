import openpyxl
from pathlib import Path
import datetime
from datetime import datetime, timedelta
data = datetime.now()
giorno = data.strftime("%Y-%m-%d 00:00:00")
xlsx_file = Path('add_folder_here', 'add_file_here')
wb_obj = openpyxl.load_workbook(xlsx_file)
ws = wb_obj.active

def convertDate(ordinal):
    epochStart = datetime(1899, 12, 31)
    if ordinal is not None:
        if ordinal >= 60:
            ordinal -= 1
        return epochStart + timedelta(days=ordinal)

row_count = ws.max_row-284
count = 0

for j in range(1, row_count, 7):
    a = ws.cell(row=j, column=1).value
    b = convertDate(a)
    if giorno == str(b):
        for x in range(j+1, 140):           # you can put ws.max_row but remember, the librarie take the last row changed,  
            if count == 6:                  # so if you had 300 rows but you edited the row 1000 too the librarie take the row 1000.
                break
            else:
                print(ws.cell(row=x,column=2).value)
                count += 1