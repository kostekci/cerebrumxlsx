from tkinter import Tk
import pandas
from openpyxl import load_workbook

d1 = {'Words': ['1']}
df1 = pandas.DataFrame(data=d1)

book = load_workbook('test.xlsx')
writer = pandas.ExcelWriter('test.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = {ws.title: ws for ws in book.worksheets}

last_text=''

while True:
    text = Tk().clipboard_get()
    if text != last_text:
        
        last_text=text
        df1['Words']=text
        for sheetname in writer.sheets:
            df1.to_excel(writer,sheet_name=sheetname, startrow=writer.sheets[sheetname].max_row, index = False,header= False)

        writer.save()