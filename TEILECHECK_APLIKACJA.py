#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
# from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import warnings
import numpy as np
import time
import xlsxwriter
import pandas as pd
# from openpyxl import load_workbook
# from openpyxl import *
from tkinter import messagebox
# import openpyxl
# from pathlib import Path
# from openpyxl import load_workbook
# from openpyxl import Workbook
# from flask-sqlalchemy import SQLAlchemy
from pathlib import Path



# warnings.simplefilter("ignore")


root= tk.Tk()
root.geometry("1100x250")
root.title('TEILECHECK VSL       Version 1.0              @ Urszula Kiryk-Kania' )  #- @Autor Urszula Kiryk-Kania (PWL-1/4)'

root.update()


val = tk.StringVar()
val_entry_box = tk.Entry(root, width=145, textvariable=val)
val_entry_box.grid(row=1, column=1)

val2 = tk.StringVar() 
val2_entry_box = tk.Entry(root, width=145, textvariable=val2)
val2_entry_box.grid(row=2, column=1)

# val3 = tk.StringVar()
# val3_entry_box = tk.Entry(root, width=145, textvariable=val3)
# val3_entry_box.grid(row=3, column=1)

# val6 = tk.StringVar()
# val6_entry_box = tk.Entry(root, width=145, textvariable=val6)
# val6_entry_box.grid(row=4, column=1)

# val5 = tk.StringVar()
# val5_entry_box = tk.Entry(root, width=145, textvariable=val5)
# val5_entry_box.grid(row=15, column=1)


# val7_text_box = tk.Text(root, height=20, width=90)
# val7_text_box.grid(row=26, column=0, padx=10, pady=10)

def UploadAction2():
    global filename2
    global df_F2
    global result_df

    filename2 = askopenfilename()

    if not filename2:
        return

    val.set(filename2)

    data2 = pd.read_excel(
        filename2,
        sheet_name='Teilecheck'
    )

    df_F = pd.DataFrame(
        data2.values[26:],
        columns=data2.iloc[26]
    )
    df_F = df_F[df_F['ZP'].astype(str).str.strip() == 'ZP7']
    jeden = "ZP"
    trzy = "Bauteile /\nParts"
    column_parts = "Teilenummer /\nPart number"
    column_type = "Bauteile /\nParts"

    rows = []

    import re

    # Numery z kropkami, np.:
    # 7E0.863.831.B
    # 03K.906.051
    #
    # Numery bez kropek, np.:
    # 03K906051
    # 05N907807BL
    wzorzec = re.compile(
        r'(?<![A-Z0-9])'
        r'(?:'
        r'(?:[A-Z0-9]{3}\.){2}[A-Z0-9]{3}'
        r'(?:\.[A-Z0-9]{1,3})?'
        r'|'
        r'[A-Z0-9]{9,12}'
        r')'
        r'(?![A-Z0-9])',
        flags=re.IGNORECASE
    )

    for _, row in df_F.dropna(
        subset=[column_parts]
    ).iterrows():

        wartosc = str(row[column_parts])

        # Usunięcie apostrofów Excela i ujednolicenie ENTER-ów
        wartosc = (
            wartosc
            .replace("\r\n", "\n")
            .replace("\r", "\n")
            .replace("'", "")
            .upper()
        )

        # findall wyszukuje wszystkie numery w całej komórce
        numery = wzorzec.findall(wartosc)

        # Usunięcie ewentualnych duplikatów z jednej komórki
        numery = list(dict.fromkeys(numery))

        for nr in numery:

            nr_czysty = nr.replace(".", "").strip()

            rows.append({
                jeden: row[jeden],
                trzy: row[trzy],
                column_type: row[column_type],
                column_parts: nr_czysty
            })

    result_df = pd.DataFrame(rows)
#     result_df = result_df.drop_duplicates(
#         subset=['ZP', 'Teilenummer /\nPart number']).reset_index(drop=True)
    # Kontrola wyniku
#     result_df.to_excel(
#         "test_plik1.xlsx",
#         index=False
#     )

    result_df[column_parts] = (
        result_df[column_parts]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.strip()
        .str.upper()
    )

    result_df['Vorseriencheck_Teilenummer'] = ''
    result_df['Begründung bei Abweichung'] = ''

    result_df['TNR_soll'] = (
        result_df[column_parts]
        .astype(str)
        .str.slice(0, 9)
    )

    return result_df

# def UploadAction2():
#     global filename2
#     global df_F2
#     global result_df
#     filename2 = askopenfilename()
#     val.set(filename2)
#     data2 = pd.read_excel(filename2, sheet_name='Teilecheck')
    
#     df_F = pd.DataFrame(data2.values[26:], columns=data2.iloc[26])
#     jeden = "ZP"
#     dwa = "nan"
#     trzy = "Bauteile /\nParts"
#     column_parts = "Teilenummer /\nPart number"
#     column_type = "Bauteile /\nParts"
#     rows = []
    
   
 
#     for _, row in df_F.dropna(subset=[column_parts]).iterrows():
        
#         wartosc = str(row[column_parts]) 
#         print(repr(wartosc))

#         for linia in wartosc.splitlines():
#             print("LINIA:", repr(linia))
#             linia = linia.replace("'", "").strip()

#             if "." in linia and any(c.isdigit() for c in linia):

#                 nr = linia.split()[0]


#                 rows.append({

#                         jeden: row[jeden],

#                         trzy: row[trzy],

#                         column_type: row[column_type],

#                         column_parts: nr.replace('.', '')

#                 })

            
            
#         numery = re.findall(
#            r'[A-Z0-9]{3}\.[A-Z0-9]{3}\.[A-Z0-9]{3}\.[A-Z0-9]{1,3}', 
#             wartosc
#         )
#         if '7E0.863.831.B' in wartosc:
 
#             with open(debug_file, "a", encoding="utf-8") as f:
 
#                 f.write("\nWARTOSC:\n")
 
#                 f.write(repr(wartosc))
 
#                 f.write("\nNUMERY:\n")

#                 f.write(str(numery))

#                 f.write("\n")
#         messagebox.showinfo("DEBUG", str(debug_file))
            
#         print(numery)

#         for nr in numery:

#                 rows.append({

#                     jeden: row[jeden],

#                     trzy: row[trzy],
    
#                     column_type: row[column_type],
 
#                     column_parts: nr.replace('.', '')
 
#                 })
    
   


# Iteracja po wierszach z niepustą kolumną z numerami części
#     for _, row in df_F.dropna(subset=[column_parts]).iterrows():
#         typ_value1 = row[jeden]  # wartość typu z tego samego wiersza
#         typ_value3 = row[trzy]
#         typ_value4 = row[column_type]  # wartość typu z tego samego wiersza
#     #     parts = row[column_parts].split(")")  # podział numerów części po spacjach
    
    
    
    
    
    
#     parts = str(row[column_parts]).split('\n')
#     print(result_df['Teilenummer /\nPart number'].head(20))
#     print(df_F['Teilenummer /\nPart number'].dropna().head(5).tolist())
# #     for part in parts:
#             rows.append({jeden: typ_value1, trzy: typ_value3, column_type: typ_value4, column_parts: part})
    # Utwórz nowy DataFrame z przypisanymi typami
# 
#     result_df.to_excel("test_plik1.xlsx", index=False)
    result_df['Teilenummer /\nPart number'] = result_df['Teilenummer /\nPart number'].astype(str).str.replace('.', '', regex=False)
    result_df['Vorseriencheck /\Teilenummer']=''
    result_df['Begründung bei /\Abweichung']=''
    result_df['TNR_soll']=result_df['Teilenummer /\nPart number'].astype(str).str.slice(0, 9)
    
    return result_df
    return val
    return filename2
  

def UploadAction():
    global filename
    global df_F
    global tnr
    filename = askopenfilename()
    val2.set(filename)
    tnr = pd.read_excel(filename)
    df_F = pd.DataFrame(tnr)
    tnr = pd.DataFrame({'TNR': tnr['Teilenummer_12']})
    tnr['TNR_cops']=tnr['TNR'].astype(str).str.slice(0, 9)
#     tnr.to_excel('test.xlsx')
    return tnr



def start(): 
    global check3
    global result_df
    global tnr
    
    check2 = pd.merge(
        result_df,
        tnr,
        how='left',
        left_on='TNR_soll',
        right_on='TNR_cops'
    )
 

    check2['exact_match'] = (
        check2['Teilenummer /\nPart number'] 
    == check2['TNR']
    )

    # dla każdego numeru sprawdź czy istnieje choć jedno idealne dopasowanie
    check2 = check2[
        (~check2.groupby('Teilenummer /\nPart number')['exact_match'].transform('any'))
        |
        (check2['exact_match'])
        ]

#     check2['check'] = np.where(
#         check2['found_exact'],
#         '',
#         'to_check'
#     )

    check2['check'] = np.where(
        check2['TNR'].isna(),
        '',
        np.where(
            check2['Teilenummer /\nPart number'] != check2['TNR'],
            'to_check',
            ''
        )
)

    print(check2[['Teilenummer /\nPart number','TNR']].head(50))

    print(len(check2))    
#     def zmien (x):
#         if x['TNR_soll'] == x['TNR_cops'] and x['Teilenummer /\nPart number'] != x['TNR']:
#              return 'to_check'
#         return ' '
#     check2['check']=check2.apply(zmien, axis=1)
#     check2['check'] = ''
 
#     check2.loc[
#          ~check2['found_exact'],
#         'check'
#     ] = 'to_check'

    def zmien2 (x):
        if  x['check'] == 'to_check':
             return 'Begründung offen'
        return ''
#     check2.to_excel('test.xlsx')

    check2['Begründung bei /\Abweichung']=check2.apply(zmien2, axis=1)
 
    check3=check2[['ZP','Bauteile /\nParts','Teilenummer /\nPart number','TNR','check','Begründung bei /\Abweichung']]
    check3=check3.rename(columns={'check':'Vorseriencheck_TNR'})
    check3=check3.rename(columns={'TNR':'TNR_COPS'})
    check3=check3.rename(columns={'Begründung bei /\Abweichung':'VSL_Begründung bei Abweichung'})
    nazwa = Path(filename).stem
    check3 = check3.drop_duplicates()#         subset=['ZP', 'Teilenummer /\nPart number']).reset_index(drop=True)
    numer = nazwa.split()[0]
    opis = nazwa.split(' Stand')[0].split(' ', 1)[1]
    DATA = time.strftime('%Y-%m-%d')
    check3.to_excel('Teilecheck_'+numer+opis+'_ZP7_VSLcheck_'+DATA+'.xlsx')
    
    
    nazwa_pliku = (
    'Teilecheck_'
    + numer
    + opis
    + '_ZP7_VSLcheck_'
    + DATA
    + '.xlsx'
    )

    with pd.ExcelWriter(
        nazwa_pliku,
        engine='xlsxwriter'
    ) as writer:

        check3.to_excel(
            writer,
            sheet_name='VSL Check',
            index=False
        )

        workbook = writer.book
        worksheet = writer.sheets['VSL Check']

        # Ukrycie siatki Excela
        worksheet.hide_gridlines(2)

        # Zamrożenie nagłówka
        worksheet.freeze_panes(1, 0)

        # Filtr w nagłówkach
        worksheet.autofilter(
            0,
            0,
            len(check3),
            len(check3.columns) - 1
        )

        # Format nagłówków
        header_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#7F1D1D',
            'border': 1,
            'border_color': '#D9D9D9',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        # Format zwykłych danych
        data_format = workbook.add_format({
            'font_color': '#000000',
            'valign': 'top',
            'text_wrap': True
        })

        # Format danych importowanych
        imported_format = workbook.add_format({
            'font_color': '#008000',
            'valign': 'top',
            'text_wrap': True
        })

        # Format ostrzeżenia to_check
        warning_format = workbook.add_format({
            'font_color': '#9C0006',
            'bg_color': '#FFC7CE',
            'bold': True,
            'align': 'center',
            'valign': 'center'
        })

        # Format uzasadnienia
        reason_format = workbook.add_format({
            'font_color': '#9C5700',
            'bg_color': '#FFEB9C',
            'valign': 'top',
            'text_wrap': True
        })

        # Format naprzemiennych wierszy
        alternate_format = workbook.add_format({
            'bg_color': '#F2F2F2',
            'valign': 'top',
            'text_wrap': True
        })

        # Ponowne zapisanie nagłówków z formatem
        for col_num, column_name in enumerate(check3.columns):
            worksheet.write(
                0,
                col_num,
                column_name,
                header_format
            )

        # Wysokość nagłówka
        worksheet.set_row(0, 35)

        # Szerokości i formaty kolumn
        worksheet.set_column('A:A', 10, data_format)       # ZP
        worksheet.set_column('B:B', 35, data_format)       # Bauteile
        worksheet.set_column('C:C', 24, imported_format)   # Teilenummer
        worksheet.set_column('D:D', 24, imported_format)   # TNR_COPS
        worksheet.set_column('E:E', 24, data_format)       # Check
        worksheet.set_column('F:F', 36, data_format)       # Uzasadnienie

        # Naprzemienne jasnoszare wiersze
        if len(check3) > 1:
            worksheet.conditional_format(
                1,
                0,
                len(check3),
                len(check3.columns) - 1,
                {
                    'type': 'formula',
                    'criteria': '=MOD(ROW(),2)=0',
                    'format': alternate_format
                }
            )

        # Czerwone wyróżnienie to_check
#         if len(check3) > 0:
#             worksheet.conditional_format(
#                 1,
#                 4,
#                 len(check3),
#                 4,
#                 {
#                     'type': 'text',
#                     'criteria': 'containing',
#                     'value': 'to_check',
#                     'format': warning_format
#                 }
#             )

#             # Żółte wyróżnienie "Begründung offen"
#             worksheet.conditional_format(
#                 1,
#                 5,
#                 len(check3),
#                 5,
#                 {
#                     'type': 'text',
#                     'criteria': 'containing',
#                     'value': 'Begründung offen',
#                     'format': reason_format
#                 }
#             )

        # Wysokość wierszy danych
        for row_num in range(1, len(check3) + 1):
            worksheet.set_row(row_num, 25)

    
    
#     messagebox.showinfo('Gotowe!', 'Raport końcowy dla auta '+numer+opis+' został zapisany w folderze:\n SHAREPOINT strona:  LOGISTYKAPRZEDSERYJNA/Dokumenty/ogólne/Systemy/TEEILECHECK/Raport')
    messagebox.showinfo('Gotowe!', 'Raport końcowy dla auta '+numer+opis+' został zapisany w lokalizacji, gdzie znajduje sie zapisana baza')

def close(): 
    root.destroy()    


# root.iconbitmap("VW.ico")


# ===== create labels ======

Plik1 = tk.Label(root, text='Plik Teilecheck: ')
Plik1.grid(row=1, column=0, sticky=tk.W, padx=(10, 5))

Plik2 = tk.Label(root, text='Plik COPS: ')
Plik2.grid(row=2, column=0, sticky=tk.W, padx=(10, 5))


# ===== create buttons ======



button1 = tk.Button(text='Pobierz', command=UploadAction2, bg='brown',activebackground='grey', fg='white', width = 10)
button1.grid(row=1, column=3)

button2 = tk.Button(text='Pobierz', command=UploadAction, bg='brown', fg='white', activebackground='grey', width = 10)
button2.grid(row=2, column=3)






name_label = tk.Label(root, text=' ')
name_label.grid(row=23, column=0,padx=5, pady=5)


name_label = tk.Label(root, text=' ')
name_label.grid(row=24, column=0,padx=5, pady=5)

info_label = tk.Label(
    root,
    text="Raport VSL generowany jest wyłącznie dla pozycji ZP7.",
    fg="dark grey",
    font=("Segoe UI", 9)
)

info_label.grid(row=22, column=1, sticky="w")

button4 = tk.Button(text='Start ', command=start, bg='brown', fg='white',activebackground='grey', width=30)
button4.grid(row=25, column=1)

name_label = tk.Label(root, text=' ')
name_label.grid(row=26, column=0,padx=5, pady=5)

button5 = tk.Button(text='Exit', command=close, bg='brown', fg='white',activebackground='grey', width=10)
button5.grid(row=26, column=3)


    
root.mainloop()


# In[ ]:




