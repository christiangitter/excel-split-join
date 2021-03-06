import glob
import os
import pandas as pd

file = input('Pfad zur Datei: ')
pth = os.path.dirname(file)
extension = os.path.splitext(file)[1]
files = glob.glob(os.path.join(pth, '*.xls*'))
newfile = os.path.join(pth, 'zusammengefuegt.xlsx')
df = pd.DataFrame()
for f in files:
    data = pd.read_excel(f)
    df = df.append(data)

df.to_excel(newfile, sheet_name='zusammengefuegt', index=False)
print('Abgeschlossen')
