#!/usr/bin/env python
# coding: utf-8

# In[7]:


import tkinter as tk
from tkinter.filedialog import askopenfilename
import os
import warnings
import numpy as np
import time
import xlsxwriter
import pandas as pd
from tkinter import messagebox
from pathlib import Path



# warnings.simplefilter("ignore")


root= tk.Tk()
root.geometry("1200x250")
root.title('TEILECHECK VSL       Version 1.0              @ Urszula Kiryk-Kania' )  #- @Autor Urszula Kiryk-Kania (PWL-1/4)'

root.update()


val = tk.StringVar()
val_entry_box = tk.Entry(root, width=150, textvariable=val)
val_entry_box.grid(row=1, column=1)

val2 = tk.StringVar() 
val2_entry_box = tk.Entry(root, width=150, textvariable=val2)
val2_entry_box.grid(row=2, column=1)



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
    print("Znalezionych numerów:", len(rows))
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

 
    result_df['Teilenummer /\nPart number'] = result_df['Teilenummer /\nPart number'].astype(str).str.replace('.', '', regex=False)
    result_df['Vorseriencheck /\Teilenummer']=''
    result_df['Begründung bei /\Abweichung']=''
    result_df['TNR_soll']=result_df['Teilenummer /\nPart number'].astype(str).str.slice(0, 9)
    
    return result_df
    
def start():
    global check3
    global result_df
    global tnr
    global filename

    # Kontrola, czy wczytano Teilecheck
    if 'result_df' not in globals() or result_df.empty:
        messagebox.showwarning(
            "Brak danych",
            "Czy plik Teilecheck, który wczytujesz, masz otwarty? Zamknij go, a dopiero potem wczytaj plik."
        )
        return

    # Kontrola, czy wczytano COPS
    if 'tnr' not in globals() or tnr.empty:
        messagebox.showwarning(
            "Brak danych",
            "Czy plik z COPS, który wczytujesz, masz otwarty? Zamknij go, a dopiero potem wczytaj plik."
        )
        return

    # Kontrola ścieżki pliku COPS
    if 'filename' not in globals() or not filename:
        messagebox.showwarning(
            "Brak pliku",
            "Nie znaleziono ścieżki pliku COPS."
        )
        return

    try:
        column_parts = "Teilenummer /\nPart number"

        # Komunikat w aplikacji
        root.config(cursor="watch")
        root.update_idletasks()

        # Połączenie danych Teilecheck z COPS
        check2 = pd.merge(
            result_df,
            tnr,
            how="left",
            left_on="TNR_soll",
            right_on="TNR_cops"
        )

        # Sprawdzenie dokładnego dopasowania numeru
        check2["exact_match"] = (
            check2[column_parts].astype(str)
            ==
            check2["TNR"].astype(str)
        )

        # Czy dla analizowanego numeru istnieje przynajmniej
        # jedno dokładne dopasowanie?
        exact_match_exists = (
            check2
            .groupby(column_parts)["exact_match"]
            .transform("any")
        )

        # Jeżeli istnieje dokładne dopasowanie,
        # usuń niedokładne warianty tego numeru
        check2 = check2[
            (~exact_match_exists)
            |
            (check2["exact_match"])
        ].copy()

        # Jeżeli COPS jest pusty, pole check pozostaje puste.
        # Jeżeli pierwsze 9 znaków pasuje, ale cały numer jest inny,
        # wpisz to_check.
        check2["check"] = np.where(
            check2["TNR"].isna(),
            "",
            np.where(
                check2[column_parts].astype(str)
                != check2["TNR"].astype(str),
                "to_check",
                ""
            )
        )

        # Uzasadnienie bez wolnego apply(axis=1)
        check2["Begründung bei /\Abweichung"] = np.where(
            check2["check"].eq("to_check"),
            "Begründung offen",
            ""
        )

        # Kolumny raportu końcowego
        check3 = check2[
            [
                "ZP",
                "Bauteile /\nParts",
                column_parts,
                "TNR",
                "check",
                "Begründung bei /\Abweichung"
            ]
        ].copy()

        # Zmiana nazw kolumn
        check3 = check3.rename(
            columns={
                "TNR": "TNR_COPS",
                "check": "Vorseriencheck_TNR",
                "Begründung bei /\Abweichung":
                    "VSL_Begründung bei Abweichung"
            }
        )

        # Usunięcie identycznych duplikatów
        check3 = (
            check3
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # Informacje z nazwy pliku COPS
        nazwa = Path(filename).stem
        elementy_nazwy = nazwa.split()

        if elementy_nazwy:
            numer = elementy_nazwy[0]
        else:
            numer = "COPS"

        if " Stand" in nazwa:
            nazwa_bez_stand = nazwa.split(" Stand")[0]
        else:
            nazwa_bez_stand = nazwa

        if " " in nazwa_bez_stand:
            opis = nazwa_bez_stand.split(" ", 1)[1].strip()
        else:
            opis = ""

        DATA = time.strftime("%Y-%m-%d")

        # Bezpieczna nazwa raportu
        elementy_pliku = [
            "Teilecheck",
            numer,
            opis,
            "ZP7",
            "VSLcheck",
            DATA
        ]

        nazwa_pliku = "_".join(
            str(element).strip().replace(" ", "_")
            for element in elementy_pliku
            if str(element).strip()
        ) + ".xlsx"

        # Zapis w folderze, w którym znajduje się plik COPS
        sciezka_wynikowa = Path(filename).parent / nazwa_pliku

        with pd.ExcelWriter(
            sciezka_wynikowa,
            engine="xlsxwriter"
        ) as writer:

            check3.to_excel(
                writer,
                sheet_name="VSL Check",
                index=False
            )

            workbook = writer.book
            worksheet = writer.sheets["VSL Check"]

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
                "bold": True,
                "font_color": "white",
                "bg_color": "#7F1D1D",
                "border": 1,
                "border_color": "#D9D9D9",
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True
            })

            # Format zwykłych danych
            data_format = workbook.add_format({
                "font_color": "#000000",
                "valign": "top",
                "text_wrap": True
            })

            # Format numerów importowanych
            imported_format = workbook.add_format({
                "font_color": "#008000",
                "valign": "top",
                "text_wrap": True
            })

            # Format naprzemiennych wierszy
            alternate_format = workbook.add_format({
                "bg_color": "#F2F2F2",
                "valign": "top",
                "text_wrap": True
            })

            # Nagłówki
            for col_num, column_name in enumerate(check3.columns):
                worksheet.write(
                    0,
                    col_num,
                    column_name,
                    header_format
                )

            worksheet.set_row(0, 35)

            # Szerokości oraz formaty kolumn
            worksheet.set_column("A:A", 10, data_format)
            worksheet.set_column("B:B", 35, data_format)
            worksheet.set_column("C:C", 24, imported_format)
            worksheet.set_column("D:D", 24, imported_format)
            worksheet.set_column("E:E", 24, data_format)
            worksheet.set_column("F:F", 36, data_format)

            # Naprzemienne jasnoszare wiersze
            if len(check3) > 1:
                worksheet.conditional_format(
                    1,
                    0,
                    len(check3),
                    len(check3.columns) - 1,
                    {
                        "type": "formula",
                        "criteria": "=MOD(ROW(),2)=0",
                        "format": alternate_format
                    }
                )

            # Wysokość wierszy
            for row_num in range(1, len(check3) + 1):
                worksheet.set_row(row_num, 25)

        messagebox.showinfo(
            "Gotowe!",
            "Raport został zapisany w lokalizacji:\n\n"
            + str(sciezka_wynikowa)
        )

    except Exception as error:
        messagebox.showerror(
            "Błąd generowania raportu",
            "Nie udało się wygenerować raportu:\n\n"
            + str(error)
        )

    finally:
        root.config(cursor="")
        root.update_idletasks() 
    
    
def reload_app():
    global result_df
    global tnr
    global filename
    global filename2
    global check3

    val.set("")
    val2.set("")

    result_df = pd.DataFrame()
    tnr = pd.DataFrame()
    check3 = pd.DataFrame()

    filename = ""
    filename2 = ""

    val_entry_box.focus_set()

    messagebox.showinfo(
        "Reload",
        "Dane zostały wyczyszczone.\n"
        "Możesz wczytać nowe pliki."
    )

 

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
button5.grid(row=30, column=2)

button_reload = tk.Button(
    text='Reload',
    command=reload_app,
    bg='grey',
    fg='white',
    activebackground='lightgrey',
    width=10
)

button_reload.grid(row=25, column=2)

    
root.mainloop()


# In[ ]:




