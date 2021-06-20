import pandas as pd
import os
from openpyxl import load_workbook
import xlsxwriter
from shutil import copyfile

file = input('Pfad zur Datei: ')
extension = os.path.splitext(file)[1]
filename = os.path.splitext(file)[0]
pth = os.path.dirname(file)
newfile = os.path.join(pth, filename+'_2'+extension)
df = pd.read_excel(file)
colpick = input('Spalte wählen, nach der geteilt werden soll : ')
cols = list(set(df[colpick].values))


def sendtofile(cols):
    for i in cols:
        df[df[colpick] == i].to_excel(
            "{}/{}.xlsx".format(pth, i), sheet_name=i, index=False)
    print('\nABGESCHLOSSEN')
    print('Vielen Dank für die Nutzung von Gittis Excel-Splitter.')
    return


def sendtosheet(cols):
    copyfile(file, newfile)
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        for myname in cols:
            mydf = df.loc[df[colpick] == myname]
            mydf.to_excel(writer, sheet_name=myname, index=False)
        writer.save()

    print('\nABGESCHLOSSEN')
    print('Vielen Dank für die Nutzung von Gittis Excel-Splitter.')
    return


print('Die Daten werden geteilt auf Basis dieser Werte {}. Im Anschluss werden {} Dateien oder Sheets erstellt. "Y" für "Weiter" "N" für "Abbruch".'.format(', '.join(cols), len(cols)))
while True:
    x = input('Weiter? (Y/N): ').lower()
    if x == 'y':
        while True:
            s = input(
                'Die Datei in "Files" oder "Sheets" teilen? (S/F): ').lower()
            if s == 'f':
                sendtofile(cols)
                break
            elif s == 's':
                sendtosheet(cols)
                break
            else:
                continue
        break
    elif x == 'n':
        print('\nVielen Dank für die Nutzung von Gittis Excel-Splitter..')
        break

    else:
        continue
