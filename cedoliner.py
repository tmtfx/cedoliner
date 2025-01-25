#    Cedoliner, elaboration, extraction and computation of pay slip parameters
#    Copyright (C) 2025  Fabio Tomat
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.

import pdfplumber
import os
import openpyxl
from openpyxl.styles import PatternFill

red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

start_row_for_year = None

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Risultati"
# Aggiungi intestazioni
ws.append(["Mese", "Pagina", "Codice", "Descrizione", "Importo"])

# Specifica la cartella contenente i PDF
cartella_pdf = "cedolini"

# Parole chiave o pattern da cercare
parole_chiave = ["0169", "0170", "0964", "0965", "0966", "0967", "0968", "0987", "0988", "0991", "0992", "0790", "0791"]
mese_anno_ref = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]#, "Tredicesima", "Quattordicesima"]

def mese_a_numero(mese):
    mesi = {
        "Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4,
        "Maggio": 5, "Giugno": 6, "Luglio": 7, "Agosto": 8,
        "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12
    }
    return mesi.get(mese, 0)

def analizza_cedolino(pdf_path, anno, parole_chiave):#, pattern_codici):
    risultati = []
    got_ref=False
    mese = None
    ispdf=True
    """questa parte ricerca il mese e l'anno del cedolino"""
    print("analisi del cedolino",pdf_path)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                testo = page.extract_text()
                if testo:
                    righe = testo.split("\n")
                    for riga in righe:
                        if not got_ref:
                            for parola in mese_anno_ref:
                                if parola in riga:
                                    n=riga.find(parola)
                                    if n!=-1:
                                        mese_anno=riga[n:]
                                        if (mese_anno.split()[1]==anno)or(mese_anno[-4:]==anno):
                                            n=len(parola)
                                            mese=mese_anno[:n]
                                            if pdf_path.lower().find(mese.lower())>-1:
                                                #anno=mese_anno[n:]
                                                print("elaborazione di",mese,anno)
                                                got_ref=True
                                                break
    except Exception as e:
            print(f"Errore nell'analisi del cedolino {pdf_path}: {e}")
            ispdf=False
    
    if ispdf:
        if not got_ref:
            print("fallback - nome mese da nome file")
            mese=os.path.basename(pdf_path).split()[0]
        """questa parte cerca i codici e le colonne in cui son scritti"""
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                testo = page.extract_text()
                tables = page.extract_tables()
                #for table in tables:
                #    print("tabella:",table,"\n")
                for table_idx, table in enumerate(tables, start=1):
                    if table:  # Se la tabella non è vuota
                        print(f"Pagina {page_num}, Tabella {table_idx}:\n")
                        
                        # Trasponi la tabella per ottenere le colonne
                        colonne = list(zip(*table))
                        # Itera sulle colonne
                        for col_idx, colonna in enumerate(colonne, start=1):
                            #print(f"Colonna {col_idx}: {colonna}\n")
                            listacolonna=[]
                            for item in colonna:
                                if item is not None:
                                    listacolonna.append(item)
                            #i = 0
                            #while i < len(colonna):
                            #    if colonna[i] is None:
                            #        colonna.pop(i)
                            #    else:
                            #        i += 1
                            if listacolonna[0]=="Trattenute":
                                trattenute=listacolonna[1].split("\n")
                                print(f"Colonna {col_idx} Trattenute: {trattenute}\n")
                            elif listacolonna[0]=="Competenze":
                                competenze=listacolonna[1].split("\n")
                                print(f"Colonna {col_idx} Competenze: {competenze}\n")
                               
                            if page_num==1:
                                if col_idx==1:
                                    i=0
                                    while i < len(listacolonna):
                                        if listacolonna[i]=="Presenze":
                                            if listacolonna[i+1]:
                                                if listacolonna[i+1].isnumeric():
                                                    presenze=listacolonna[i+1]
                                                    print(f"Presenze: {presenze}\n")
                                                else:
                                                    presenze=listacolonna[i+1]
                                                    print(f"Presenze non numeriche: {presenze}\n")
                                            else:
                                                print("Nessuna presenza\n")
                                                presenze=0
                                        i+=1
                                else:
                                    if listacolonna[0]=="Ferie":
                                        if listacolonna[1]:
                                            if listacolonna[1].isnumeric():
                                                ferie=listacolonna[1]
                                                print(f"Ferie: {ferie}\n")
                                            else:
                                                ferie=listacolonna[1]
                                                print(f"Ferie non numeriche: {ferie}\n")
                                        else:
                                            print("Nessun giorno di ferie\n")
                                            ferie=0
                if testo:
                    righe = testo.split("\n")
                    for riga in righe:
                        if any(parola in riga for parola in parole_chiave):
                            for parola in parole_chiave:
                                if parola.lower() in riga.lower():
                                    print(f"riga: {riga}\n")
                                    #recupero descrizione
                                    ns=riga.lower().find(parola.lower())+len(parola)
                                    ne=riga.find(" X ")
                                    if ne==-1:
                                        ne=len(riga)-len(parola)
                                    descrizione=riga[:ne][ns:]
                                    #print(descrizione) 
                                    #tables = page.extract_tables()
                                    if descrizione!="":
                                        valori=riga.split()
                                        #print("valori:",valori)
                                        if valori[-2].find("-")>-1:
                                            #print(f"trovata trattenuta {valori[-1]}")
                                            valore="-"+valori[-1]
                                        else:
                                            #print(f"trovata competenza {valori[-1]}")
                                            valore=valori[-1]
                                        risultati.append((page_num, mese, parola, valore,descrizione))
                                        ##print("descrizione:",descrizione)
                                        #for table in tables:
                                        #    #print(table)
                                        #    for row in table:
                                        #        #print(row)
                                        #        i = 0
                                        #        while i < len(row):
                                        #            if row[i] is None:
                                        #                row.pop(i)
                                        #            else:
                                        #                i += 1
                                        #        if any(parola.lower() in cell.lower() for cell in row if cell):
                                        #            #print("trovato",parola,"in",row)
                                        #            #print("per il mese",mese,anno)
                                        #            #print("array di row:",row)
                                        #            if row[-1]:
                                        #                valore = row[-1]
                                        #                #print(row[-1])
                                        #                if riga.split()[-1] in valore.split("\n"):
                                        #                    print("caso 1.1",riga.split()[-1])
                                        #                    #print("aggiungo "+riga.split()[-1])
                                        #                else:
                                        #                    risultati.append((page_num, mese, parola, "-"+riga.split()[-1],descrizione))
                                        #                    print("caso 1.2",riga.split()[-1])
                                        #                    #print("aggiungo -"+riga.split()[-1])
                                        #            elif row[-2]:
                                        #                valore = row[-2]
                                        #                #print(row[-2])
                                        #                if riga.split()[-1] in valore.split("\n"):
                                        #                    risultati.append((page_num, mese, parola, "-"+riga.split()[-1],descrizione))
                                        #                    print("caso 2.1",riga.split()[-1])
                                        #                    #print("aggiungo -"+riga.split()[-1])
                                        #                else:
                                        #                    risultati.append((page_num, mese, parola, riga.split()[-1],descrizione))
                                        #                    print("caso 2.2",riga.split()[-1])
                                        #                    #print("aggiungo "+riga.split()[-1])
                                        #            else:
                                        #                print("ERRORE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        return risultati
    else:
        return []


# Ottieni la lista di tutti i file PDF nella cartella
pdf_files = []
risultati_cartella = []
for root, dirs, files in os.walk(cartella_pdf):
    for file in files:
        if file.lower().endswith('.pdf'):
            pdf_files.append((os.path.join(root, file), os.path.basename(root)))
    if os.path.basename(root).isnumeric():
        ws.append(["Anno:", os.path.basename(root)])
        cell = ws.cell(row=ws.max_row, column=1)
        cell.font = openpyxl.styles.Font(size=14, bold=True)
        cell = ws.cell(row=ws.max_row, column=2)
        cell.font = openpyxl.styles.Font(size=14, bold=True)
    conta=0
    for pdf_path in pdf_files:
        conta+=1
        risultati = analizza_cedolino(pdf_path[0], pdf_path[1], parole_chiave)#, pattern_codici)
        if len(risultati)==0:
            print(f"ATTENZIONE: il file [{pdf_path[0]}] non ha prodotto risultati")
        risultati_cartella.extend(risultati)
    if conta<12:
        print("ATTENZIONE: sono stati elaborati meno di 12 file")
    elif conta==12:
        print("ATTENZIONE: sono stati elaborati 12 files, verifica che:\nla tredicesima sia segnata all'interno di uno dei file o che ci siano tutti i file")
    else:
        print("file elaborati:",conta)
    risultati_ordinati = sorted(risultati_cartella, key=lambda x: mese_a_numero(x[1]))




    for pagina, mese, parola,valore,descrizione in risultati_ordinati:
        if start_row_for_year is None:
            start_row_for_year = ws.max_row + 1  # Set the starting row for the current year
        ws.append([mese, pagina, parola,descrizione, float(valore.replace(",",".")),"",pdf_path[1]])
        cell = ws.cell(row=ws.max_row,column=4)
        if str(valore).startswith("-"):
            cell.fill = red_fill
    #ws.append(["", "", "", f"Totale {pdf_path[1]}", f"=SUM(E2:E{ws.max_row})"])
    if start_row_for_year is not None:
        #ws[f"E{ws.max_row + 1}"] = f"=SUM(E{start_row_for_year}:E{ws.max_row})"
        ws.append(["","","",f"Totale {pdf_path[1]}", f"=SUM(E{start_row_for_year}:E{ws.max_row})"])
    pdf_files.clear()# = []
    risultati_cartella.clear()  # Clear the results for the next year
    start_row_for_year = None

wb.save("risultati_cedolini.xlsx")

print("File Excel creato con successo: risultati_cedolini.xlsx")
