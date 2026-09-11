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



warnings.simplefilter("ignore")


root= tk.Tk()
root.geometry("1500x200")
root.title('RA Fruewarnsystem Werk 35                      Version 1.1\n                  @Autor Urszula Kiryk-Kania' )  #- @Autor Urszula Kiryk-Kania (PWL-1/4)'

root.update()


val = tk.StringVar()
val_entry_box = tk.Entry(root, width=145, textvariable=val)
val_entry_box.grid(row=1, column=1)

val2 = tk.StringVar() 
val2_entry_box = tk.Entry(root, width=145, textvariable=val2)
val2_entry_box.grid(row=2, column=1)

val3 = tk.StringVar()
val3_entry_box = tk.Entry(root, width=145, textvariable=val3)
val3_entry_box.grid(row=3, column=1)

val4 = tk.StringVar()
val4_entry_box = tk.Entry(root, width=145, textvariable=val4)
val4_entry_box.grid(row=4, column=1)



def UploadAction():
    global filename
    global df_F
    filename = askopenfilename()
    val.set(filename)
    data1 = pd.read_excel(filename)
    df_F = pd.DataFrame(data1.values[1:], columns=data1.iloc[0])
    df_F.header=0
    df_F.drop_duplicates(subset=['SG-TNR','Lieferscheinnummer','Lieferscheinpositionsnummer '], keep='first', inplace=True)
    button1.config(bg="#77dd77")
    button1['text'] = 'Pobrano'
    df_F.sort_values(by=['Lieferscheinnummer'])
    df_F.drop_duplicates(subset=['SG-TNR','Lieferscheinnummer'], keep='first', inplace=True)
    return val
    return filename

def UploadAction2():
    global df8
    global df
    global df_F
    global dff
    global dfsw
    filename2 = askopenfilename()
    val2.set(filename2)
    data = pd.read_excel(filename2)
    df =pd.DataFrame(data.values[0:], columns=data.iloc[0])
    df.header=0
    df=df.rename(columns={'2KA Status':'Status'})
    df=df[['SG-TNR', 'HW-V', 'SW-V', 'Reifegrad', 'Status']]
    df=df[df['Status'].isin(['Verbaufähig','Verbaufähig für P und KD','Verbaufähig für P und KD mit Einschränkungen', 'Bestellstopp für P'])]
    df_4 = (
        df
        .groupby('SG-TNR')
        .agg({'HW-V': lambda x: list(x)})
        .reset_index()
    )
    dff=pd.concat([df_4['SG-TNR'], pd.DataFrame(df_4['HW-V'].values.tolist()).add_prefix('HW-V_')], axis=1)
    dfa=df_4
    HW= pd.merge(df_4, dff, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
    HW2= pd.merge(df, HW, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
    df_5 = (
        df
        .groupby('SG-TNR')
        .agg({'SW-V': lambda x: list(x)})
        .reset_index()
        )
    dfsw=pd.concat([df_5['SG-TNR'], pd.DataFrame(df_5['SW-V'].values.tolist()).add_prefix('SW-V_')], axis=1)
    df2=pd.merge(df_5, dfsw, how='inner', left_on=['SG-TNR'], right_on=['SG-TNR'])
    df2= pd.merge(HW, df2, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
    df8= pd.merge(df2, df, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
    df8.drop_duplicates(subset=['SG-TNR'], keep='first', inplace=True)
    def zmien (x):
        if x['Status'] == 'Verbaufähig' or x['Status'] == 'Verbaufähig für P und KD' or x['Status'] == 'Verbaufähig für P und KD mit Einschränkungen' or x['Status'] ==  'Bestellstopp für P':
            return 'FREI'
        return ''
    df8['Freigabe']=df8.apply(zmien, axis=1)
    button2.config(bg="#77dd77")
    button2['text'] = 'Pobrano'
    return df8
    return dff
    return dfsw
    
def UploadAction3():
    global df5
#     filename3 = "X:/Groups/PWL/PWL1/PWL14/08_Systemowe/8_bazy danych/Stale_dane/Stałe dane wszystkie.xlsx"
    onedrive = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
    filename3 = Path(onedrive) / "08_Systemowe" / "8_bazy danych" / "Stale_dane" / "Stałe dane wszystkie.xlsx"
    val3.set(filename3)
    data3 = pd.read_excel(filename3)
    df5 =pd.DataFrame(data3)
    df5 = df5[df5['Werk Akt.'] == 35] 
    df5=df5.rename(columns={'Sachnummer Akt.':'SG-TNR'})
    str(df5['SG-TNR'])
    df5['SG-TNR']=df5['SG-TNR'].str.replace(' ', '')
    button3.config(bg="#77dd77")
    button3['text'] = 'Załadowano'
    return val3
    return filename3
    return df5

def mapp():
    global map1
    filename4 = askopenfilename()
    val4.set(filename4)
    mapa = pd.read_excel(filename4)
    map1=pd.DataFrame(mapa.values[1:],columns=mapa.iloc[0])

#     map1=map1.rename(columns={'ZSB-TNR':'SG-TNR'})
#     str(map1['ZSB-TNR'])
#     map1['ZSB-TNR']=m1p1['ZSB-TNR'].str.replace(' ', '')
    button6.config(bg="#77dd77")
    
    
    button6['text'] = 'Pobrano'
    return map1

        
def close(): 
    root.destroy() 



def start(): 
    global cwd
    global Raport
    global map1
    
    Raport_35 = pd.merge(df_F, df8, how='left', left_on=['SG-TNR'], right_on=['SG-TNR'])
    Raport_ZSB=Raport_35[['SG-TNR','Farbkennzeichen','HW-V','SW-V','Status']]
    Raport_35['SG-TNR_bezFKZ']=Raport_35['SG-TNR']
    Raport_35['SG-TNR']=Raport_35['SG-TNR'].astype(str)+Raport_35['Farbkennzeichen'].astype(str)
    Raport_35['SG-TNR']=Raport_35['SG-TNR'].str.replace('nan', '')
    Raport_WK35 = pd.merge(Raport_35, df5, how='left', left_on=['SG-TNR'], right_on=['SG-TNR'])
    Raport_WK35 = pd.merge(Raport_WK35, dff, how='left', left_on=['SG-TNR'], right_on=['SG-TNR'])
    Raport_WK35 = pd.merge(Raport_WK35, dfsw, how='left', left_on=['SG-TNR'], right_on=['SG-TNR'])
    ha=Raport_WK35.loc[:,Raport_WK35.columns.str.startswith('HW')]
    ha=ha.rename(columns={'HW-V_x':'HW-V_SOLL','HW-V':'HW-V_MAT'})
    del ha['HW-V_SOLL']
    del ha['HW-V_y']
#     Raport_WK35=Raport_WK35.rename(columns={'HW-V_x':'HW-V_SOLL','SW-V_x':'SW-V_SOLL','HW-V':'HW-V_MAT','SW-V':'SW-V_MAT'})
#     Raport_WK35=pd.merge(Raport_WK35_HW, df8, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
#     ha=Raport_WK35[['SG-TNR','HW-V_MAT']]
#     had=ha.iloc[:,1:]
    cols = ha.columns
    dla=ha[cols].eq(ha['HW-V_MAT'], axis=0)
    dla.astype(bool).mask(dla.isna())
    dla_new=dla.iloc[:,1:]
    dla_new=dla_new*1
    results = dla_new.sum(axis=1)
    results=pd.concat([Raport_WK35.iloc[:,2:3],results], axis=1)
    results.columns=[*results.columns[:-1],'Check_HW']
    sa=Raport_WK35.loc[:,Raport_WK35.columns.str.startswith('SW')]
    sa=sa.rename(columns={'SW-V_x':'SW-V_SOLL','SW-V':'SW-V_MAT'})
    del sa['SW-V_SOLL']
    sa = sa.loc[:,~sa.columns.str.contains('_x', case=True)]
#     Raport_WK35=Raport_WK35.rename(columns={'HW-V_x':'HW-V_SOLL','SW-V_x':'SW-V_SOLL','HW-V':'HW-V_MAT','SW-V':'SW-V_MAT'})
#     Raport_WK35=pd.merge(Raport_WK35_HW, df8, how='outer', left_on=['SG-TNR'], right_on=['SG-TNR'])
#     ha=Raport_WK35[['SG-TNR','HW-V_MAT']]
#     had=ha.iloc[:,1:]
    cols = sa.columns
    dlasw=sa[cols].eq(sa['SW-V_MAT'], axis=0)
    dlasw.astype(bool).mask(dlasw.isna())
    dla_new2=dlasw.iloc[:,1:]
    dla_new2=dla_new2*1
    results2 = dla_new2.sum(axis=1)
    results2=pd.concat([Raport_WK35.iloc[:,2:3],results2], axis=1)
    results2.columns=[*results2.columns[:-1],'Check_SW']
#     results = pd.DataFrame(data=results.values, columns=['SG-TNR','Farbkennzeichen','Check_HW'])
    
    
    
    alla=pd.concat([Raport_WK35,results.iloc[:,1:]], axis=1)
    alla=pd.concat([alla,results2.iloc[:,1:]], axis=1)
    
#     def czysc_HW(x):
#         if x['Check_HW'] > 0:
#             return True
#         return False
#     alla['Check_HW'] = alla.apply(czysc_HW, axis=1)


#     def czysc_SW(x):
#         if x['Check_SW'] > 0:
#             return True
#         return False
#     alla['Check_SW'] = alla.apply(czysc_SW, axis=1)

    alla=alla[['DA','SG-TNR','Sachnummer Bezeichnung Akt.','Lieferantennummer','Lieferantenindex','Lieferantenname','Disponent Akt.', 'Lieferscheinnummer', 'Lieferscheindatum','HW-V','SW-V', 'Teilegenerationsstand', 'Soll-Eintreffdatum', 'Lager','Werk','Freigabe','Status','HW-V_x','SW-V_x','Check_HW','Check_SW','SG-TNR_bezFKZ']]
    alla=alla.rename(columns={'Sachnummer Bezeichnung Akt.':'Bennenung','Disponent Akt.': 'DISPO','Lieferscheinnummer': 'LieferscheinNR','Teilegenerationsstand': 'GS','Lieferantennummer': 'LNR','Lieferantenindex': 'LNRindex','HW-V': 'HW-V_MAT','SW-V': 'SW-V_MAT','HW-V_x': 'HW-V_SOLL','SW-V_x': 'SW-V_SOLL'})
    alla['LNR']=alla['LNR'].astype('str').str.zfill(8)
    def czysc_LNR(x):
        if x['LNR'] == '00000nan':
            return ''
        return x['LNR']
    alla['LNR'] = alla.apply(czysc_LNR, axis=1)
    
#     def check_SW_HW(x):
#         if x['Check_HW'] == False or x['Check_SW'] == False:
#             return False
#         return True
#     alla['Check_HW_SW'] = alla.apply(check_SW_HW, axis=1)

    
    alla=alla[['SG-TNR', 'Bennenung', 'LNR', 'LNRindex', 'Lieferantenname', 'DISPO', 'LieferscheinNR', 'Lieferscheindatum', 'HW-V_MAT', 'SW-V_MAT', 'GS', 'Soll-Eintreffdatum', 'Lager', 'Werk', 'Freigabe','Status', 'HW-V_SOLL', 'SW-V_SOLL', 'Check_HW', 'Check_SW','SG-TNR_bezFKZ']]
    alla=alla.rename(columns={'Status DL24':'Freigabestatus'})

    alla.dropna(subset=['LieferscheinNR'], inplace=True)
    DZISIAJ = time.strftime('%Y-%m-%d__%H_%M_%S')
#     alla['Check_HW_SW']=alla['Check_HW']+alla['Check_SW']
    
#     def czysc_SW_HW(x):
#         if x['Check_HW_SW'] > 0:
#             return True
#         return False
#     alla['Check_HW_SW'] = alla.apply(czysc_SW_HW, axis=1)
    
    

    
    mapping = pd.merge(alla, map1, how='left', left_on=['SG-TNR_bezFKZ'], right_on=['ZSB-TNR'])
    mapping = pd.merge(mapping, df8, how='left', left_on=['SG-TNR_y'], right_on=['SG-TNR'])
    mapping2=mapping[mapping['ZSB-TNR'].isin(['HW-V_SOLL','SW-V_SOLL','Status', 'Freigabe'])]
    
    mapping['ZSB_HW'] =  mapping['SG-TNR'].astype(str)+': '+mapping['HW-V_x'].astype(str)
    mapping['ZSB_SW'] =  mapping['SG-TNR'].astype(str)+': '+mapping['SW-V_x'].astype(str)

#     mapping[mapping['Status_x'].isna()] = np.nan
    mapping['Status_x_null']=mapping['Status_x'].isnull()
    def ZSB_Status(x):
        if x['Status_x_null'] == True:
             return x['Status_y']
        return x['Status_x']
    mapping['Status_x'] = mapping.apply(ZSB_Status, axis=1)
    
    mapping['Freigabe_x_null']=mapping['Freigabe_x'].isnull()
    def ZSB_Freigabe(x):
        if x['Freigabe_x_null'] == True:
             return x['Freigabe_y']
        return x['Freigabe_x']
    mapping['Freigabe_x'] = mapping.apply(ZSB_Freigabe, axis=1)
    
    mapping['HW_V_SOLL_null']=mapping['HW-V_SOLL'].isnull()
    def ZSB_HW(x):
        if x['HW_V_SOLL_null'] == True:
             return x['ZSB_HW']
        return x['HW-V_SOLL']
    mapping['HW-V_SOLL'] = mapping.apply(ZSB_HW, axis=1)
    
    mapping['SG-TNR_y']=mapping['SG-TNR_y'].isnull()
    def ZSB_Y_N(x):
        if x['SG-TNR_y'] > 0:
            return 'N'
        return 'Y'
    mapping['ZSB'] = mapping.apply(ZSB_Y_N, axis=1)
    
    mapping['SW_V_SOLL_null']=mapping['SW-V_SOLL'].isnull()
    def ZSB_SW(x):
        if x['SW_V_SOLL_null'] == True:
             return x['ZSB_SW']
        return x['SW-V_SOLL']
    mapping['SW-V_SOLL'] = mapping.apply(ZSB_SW, axis=1)
    

    
    
    #check dla ZSB SW
    ma=mapping.loc[:,mapping.columns.str.startswith('SW')]
    del ma['SW-V_SOLL']
    del ma['SW-V_x']
    del ma['SW-V_y']
    del ma['SW_V_SOLL_null']
    
    colsma = ma.columns
    zsbsw=ma[colsma].eq(ma['SW-V_MAT'], axis=0)
    zsbsw.astype(bool).mask(zsbsw.isna())
    dla_new3=zsbsw.iloc[:,1:]
    dla_new3=dla_new3*1
    results3 = dla_new3.sum(axis=1)
    results3=pd.concat([mapping.iloc[:,0:1],results3], axis=1)
    results3.columns=[*results3.columns[:-1],'Check_ZSBSW']
    resultszsbsw = pd.DataFrame(data=results3.values, columns=['SG-TNR','Check_ZSBSW'])
    mapping=pd.concat([mapping,resultszsbsw.iloc[:,1:]], axis=1)

    def ZSB_SW_False(x):
        if x['Check_SW'] == False:
             return x['Check_ZSBSW']
        return x['Check_SW']
    mapping['Check_SW'] = mapping.apply(ZSB_SW_False, axis=1)
    
    
        
    def czysc_SW(x):
        if x['Check_SW'] > 0:
            return True
        return False
    mapping['Check_SW'] = mapping.apply(czysc_SW, axis=1)
    
    
   #check dla ZSB HW
    mah=mapping.loc[:,mapping.columns.str.startswith('HW')]
    del mah['HW-V_SOLL']
    del mah['HW-V_x']
    del mah['HW-V_y']
    del mah['HW_V_SOLL_null']
    
    colsmah = mah.columns
    zsbhw=mah[colsmah].eq(mah['HW-V_MAT'], axis=0)
    zsbhw.astype(bool).mask(zsbhw.isna())
    dla_new5=zsbhw.iloc[:,1:]
    dla_new5=dla_new5*1
    results4 = dla_new5.sum(axis=1)
    results4=pd.concat([mapping.iloc[:,0:1],results4], axis=1)
    results4.columns=[*results4.columns[:-1],'Check_ZSBHW']
    resultszsbhw = pd.DataFrame(data=results4.values, columns=['SG-TNR','Check_ZSBHW'])
    mapping=pd.concat([mapping,resultszsbhw.iloc[:,1:]], axis=1)

    def ZSB_HW_False(x):
        if x['Check_HW'] == False:
             return x['Check_ZSBHW']
        return x['Check_HW']
    mapping['Check_HW'] = mapping.apply(ZSB_HW_False, axis=1)
    
    def czysc_HW(x):
        if x['Check_HW'] > 0:
            return True
        return False
    mapping['Check_HW'] = mapping.apply(czysc_HW, axis=1)
    
    def czysc_HW(x):
        if x['Check_HW'] > 0:
            return True
        return False
    mapping['Check_HW'] = mapping.apply(czysc_HW, axis=1)
    
    
    def check_SW_HW(x):
        if x['Check_HW'] == False or x['Check_SW'] == False:
            return False
        return True
    mapping['Check_HW_SW'] = mapping.apply(check_SW_HW, axis=1)
    
    mapping=mapping.rename(columns={'Freigabe_x':'Freigabe'})
    mapping=mapping.rename(columns={'Status_x':'Status'})
    mapping=mapping.rename(columns={'SG-TNR':'Sterownik'})
    
    mapping=mapping.rename(columns={'SG-TNR_x':'TNR'})
#     mapping=mapping.rename(columns={'ZSB':'STG'})
  
    def czysc_nan_zsbhw(x):
        if x['HW-V_SOLL'] == 'nan: nan':
            return ' '
        return x['HW-V_SOLL']
    mapping['HW-V_SOLL'] = mapping.apply(czysc_nan_zsbhw, axis=1)
    
    def czysc_nan_zsbsw(x):
        if x['SW-V_SOLL'] == 'nan: nan':
            return ' '
        return x['SW-V_SOLL']
    mapping['SW-V_SOLL'] = mapping.apply(czysc_nan_zsbsw, axis=1)

    
    mapping=mapping[['TNR', 'Bennenung', 'LNR', 'LNRindex', 'Lieferantenname', 'DISPO', 'LieferscheinNR', 'Lieferscheindatum', 'HW-V_MAT', 'SW-V_MAT', 'GS', 'Soll-Eintreffdatum', 'Lager', 'Werk', 'Freigabe','Status','ZSB', 'HW-V_SOLL', 'SW-V_SOLL', 'Check_HW', 'Check_SW','Check_HW_SW']]

    
    
    DZISIAJ = time.strftime('%Y-%m-%d__%H_%M_%S')
    with pd.ExcelWriter('Raport_Fruehwarnsystem_WK35_'+DZISIAJ+'.xlsx') as writer:
        mapping.to_excel(writer, sheet_name = "Raport", index=False)# header=None)        # , header=None)
        workbook = writer.book
        worksheet = writer.sheets['Raport']
        worksheet.freeze_panes(1, 1)
        worksheet.set_zoom(73)
        #adding autofilter
        max_row, max_col = alla.shape
        worksheet.autofilter(0,0,max_row, max_col)
        for col in ('Q'):
            writer.sheets['Raport'].set_column(col+':'+col, 3)
        for col in ('C','D','I','J', 'K', 'M', 'N','P','O'):
            writer.sheets['Raport'].set_column(col+':'+col, 10)
        for col in ('F','G'):
            writer.sheets['Raport'].set_column(col+':'+col, 15)
        for col in ('H','L'):
            writer.sheets['Raport'].set_column(col+':'+col, 19)
        for col in ('A','B','E'):
            writer.sheets['Raport'].set_column(col+':'+col, 20)
        for col in ('R','S'):
            writer.sheets['Raport'].set_column(col+':'+col, 23)
        format1 = workbook.add_format({'bg_color': "#FFC7CE", 'font_color': '#9C0006'})  #red
        format2 = workbook.add_format({'bg_color': "#C6EFCE", 'font_color': "#006100"})    #green
        worksheet.conditional_format('O1:O1048576', {'type': 'text','criteria': 'containing','value': 'F', 'format': format2})
        worksheet.conditional_format('V1:V1048576', {'type': 'text','criteria': 'containing','value': 'F', 'format': format1})
        worksheet.conditional_format('V1:V1048576', {'type': 'text', 'criteria': 'containing','value': 'P', 'format': format2})
        writer.sheets['Raport'].set_column(col+':'+col, 40)
        writer.save()
        button4.config(bg="#77dd77")
        button4['text'] = 'Gotowe'
    cwd = os.getcwd() 
    messagebox.showinfo('Gotowe!', 'Raport końcowy został zapisany w folderze:\n' + cwd)
      

# root.iconbitmap("VW.ico")


# ===== create labels ======

Plik1 = tk.Label(root, text='Werk 35 -Pobierz plik Anzeige der ankommenden SG-Versionen: ')
Plik1.grid(row=1, column=0, sticky=tk.W, padx=(10, 5))

Plik2 = tk.Label(root, text='Werk 35 -plik Steuergeräteversionsliste: ')
Plik2.grid(row=2, column=0, sticky=tk.W, padx=(10, 5))


Plik3 = tk.Label(root, text='Werk 35 -Załaduj plik z DISPO: ')
Plik3.grid(row=3, column=0, sticky=tk.W, padx=(10, 5))

Plik4 = tk.Label(root, text='Werk 35 -Pobierz ZSB mapping: ')
Plik4.grid(row=4, column=0, sticky=tk.W, padx=(10, 5))




label10 = tk.Label(text=' ')
label10.grid(row=5, column=1,padx=2, pady=2)


label14 = tk.Label(text=' ')
label14.grid(row=6, column=2,padx=5, pady=5)







# ===== create buttons ======



button1 = tk.Button(text='Pobierz', command=UploadAction, bg='brown',activebackground='grey', fg='white', width = 30)
button1.grid(row=1, column=2)

button2 = tk.Button(text='Pobierz', command=UploadAction2, bg='brown', fg='white', activebackground='grey', width = 30)
button2.grid(row=2, column=2)

button3 = tk.Button(text='Załaduj', command=UploadAction3, bg='brown', fg='white',activebackground='grey', width=30)
button3.grid(row=3, column=2)

button6 = tk.Button(text='Pobierz', command=mapp, bg='brown', fg='white',activebackground='grey', width=30)
button6.grid(row=4, column=2)


button4 = tk.Button(text='Start ', command=start, bg='brown', fg='white',activebackground='grey', width=30)
button4.grid(row=6, column=1)


button5 = tk.Button(text='Exit', command=close, bg='brown', fg='white',activebackground='grey', width=10)
button5.grid(row=7, column=2)


    
root.mainloop()


# In[ ]:




