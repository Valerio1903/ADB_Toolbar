# -*- coding: utf-8 -*-
""" Verifica che gli elementi esportati in IFC siano gli stessi presenti nella vista di esportazione del modello.  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica\nExport IFC'
import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
from pyrevit import forms, script
import sys
import clr
import System
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import math
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
output = pyrevit.output.get_output()
aview = doc.ActiveView
##############################################################

def ConvertiUnita(valore):
    return UnitUtils.Convert(valore, UnitTypeId.Feet, UnitTypeId.Meters)

######################################################################################

# script.py - Lanciato da PyRevit
import os
import sys
import pyrevit

# CREAZIONE LISTE DI OUTPUT DATA
IFC_MATCH_CSV_OUTPUT = []
IFC_MATCH_CSV_OUTPUT.append(["Nome Elemento","Ifc Class","GUID","ID Elemento","Presenza Nativo","Presenza IFC","Stato"])

# Trova il percorso della cartella corrente (dove si trova anche lettore.py)
current_folder = os.path.dirname(__file__)
lettore_script = os.path.join(current_folder, "lettore.py")

# Percorso all'interprete Python 3 (adatta in base alla tua installazione)
python_exe = r"C:\Users\2Dto6D\AppData\Local\Programs\Python\Python311\python.exe"  # <-- Sostituisci con il tuo path se necessario

# Usa System.Diagnostics.Process per eseguire il comando e aspettare che finisca
try:
    process = System.Diagnostics.Process()
    process.StartInfo.FileName = python_exe
    process.StartInfo.Arguments = lettore_script
    process.StartInfo.UseShellExecute = False
    process.Start()
    process.WaitForExit()  # Aspetta che il processo finisca
    #output.print_md("# lettore.py avviato con successo")

except Exception as e:
    output = pyrevit.output.get_output()
    output.print_md("## ❌ Errore nell'esecuzione dello script lettore.py")
    output.print_md(str(e))

# Ora chiama la funzione leggi_csv
def leggi_csv():
    # Definisci il percorso del file CSV (nella cartella corrente)
    csv_path = os.path.join(current_folder, "risultati_ifc.csv")
    
    try:
        # Apri il file CSV e leggi le righe
        with open(csv_path, mode='r') as file:
            reader = csv.reader(file)
            next(reader)  # Salta la prima riga (intestazioni)
            dati = [riga for riga in reader]  # Leggi tutte le righe successive
        return dati
    
    except OSError as e:
        output.print_md("## ❌ Errore nella lettura del file CSV: {}".format(str(e)))
        return []

def EstraiInfoOggetto(oggetto):
    # SE L'OGGETTO E' UNA FAMIGLIA CARICABILE MI PRENDE FAMIGLIA E TIPO
    if isinstance(oggetto,FamilyInstance):
        return [oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
    # SE L'OGGETTO E' UNA FAMIGLIA DI SISTEMA MI PRENDE SOLO IL TIPO
    else:
        try:
            #return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            return [False,oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
        except:
            try:
                return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
            except:
                return "NON TROVATO"


GuidFile, ElementIdFile, IfcClass, NomeFile = [], [], [], []
GuidNativo, ElementIdNativo, IfcClassNativo, NomeNativo = [], [], [], []
dati_csv = leggi_csv()

if dati_csv:
    for riga in dati_csv:
        nome, tipo, guid, elementid  = riga
        #print("Tipo: {}, GUID: {}".format(tipo, guid))
        NomeFile.append(nome)
        IfcClass.append(tipo)
        GuidFile.append(guid)
        ElementIdFile.append(elementid)

else:
    print("Nessun dato trovato nel CSV.")


# Estrazione elementi in vista Revit - VALORIZZATI CON IFC EXPORT ELEMENT AS
Elementi = FilteredElementCollector(doc,aview.Id).ToElements()
for Elemento in Elementi:

    if Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS) and Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsString() != None:
        InfoNome = EstraiInfoOggetto(Elemento)

        if InfoNome[0] == False:
            NomeNativo.append(InfoNome[1])
        else:
            NomeNativo.append("-".join([InfoNome[0],InfoNome[1]]))
        
        GuidNativo.append(Elemento.get_Parameter(BuiltInParameter.IFC_GUID).AsValueString())
        ElementIdNativo.append(Elemento.Id)
        IfcClassNativo.append(Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsValueString())
        #print(Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsString())



DataTableIFC = []
DataTableNativo = []
CONTROLLO = 0

output = pyrevit.output.get_output()
output.print_md("# Verifica Elementi Corrispondenza IFC - Nativo")
output.print_md("---")

#! VERIFICA ELEMENTI PRESENTI IN IFC E NON IN NATIVO
for nome, tipo, guid, elementid in zip(NomeFile,IfcClass,GuidFile,ElementIdFile):
    VALUE = 1
    IFC = 1
    NATIVO = 1
    if guid not in GuidNativo:
        DataTableIFC.append([nome, tipo, guid, elementid, ":white_heavy_check_mark:", ":cross_mark:"])
        CONTROLLO = 1
        VALUE = 0
        NATIVO = 0
        IFC = 1
    IFC_MATCH_CSV_OUTPUT.append([nome, tipo, guid, elementid,NATIVO,IFC, VALUE])
    
#! VERIFICA ELEMENTI PRESENTI IN NATIVO E NON IN IFC
for nome, tipo, guidnat,elementid in zip(NomeNativo,IfcClassNativo,GuidNativo,ElementIdNativo):
    VALUE = 1
    if guidnat not in GuidFile:
        DataTableNativo.append([nome, tipo, guidnat, output.linkify(elementid), ":cross_mark:", ":white_heavy_check_mark:"])
        CONTROLLO = 2
        VALUE = 0
        NATIVO = 1
        IFC = 0
    IFC_MATCH_CSV_OUTPUT.append([nome, tipo, guid, elementid,NATIVO,IFC,VALUE])

#! PRINT TABELLA IN CASO DI VERIFICA FALLITA
if CONTROLLO != 0:

    output.print_md("###Elementi presenti in IFC e non in Nativo")
    try:
        output.print_table(table_data = DataTableIFC, columns = ["Nome Elemento", "Ifc Class", "GUID","ID Elemento", "Presente in IFC","Presente in Nativo"])
        output.print_md("---")
    except:
        output.print_md("###Niente da segnalare.")

    output.print_md("###Elementi presenti Nativo e non in IFC")
    try:
        output.print_table(table_data = DataTableNativo, columns = ["Nome Elemento","Ifc Class", "GUID","ID Elemento", "Presente in IFC","Presente in Nativo"])
    except:
        output.print_md("###Niente da segnalare.") 

else:
    output.print_md("## ✅ Tutti gli elementi sono presenti in IFC e nel modello.")

"""
#! VERIFICA ELEMENTI SENZA IFC EXPORT ELEMENT AS
Counter = 0
for Elemento in Elementi:
    if Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS) and Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsString() == None and hasattr(Elemento, "SystemType"):
        Counter +=1
if Counter != 0:
    output.print_md("<b style=color:orange;'>Sono stati rilevati elementi senza IFC Export As, verificare tramite l'apposito tastino.</b>".format(Counter))
"""


###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si", "No"]
Scelta = pyrevit.forms.CommandSwitchWindow.show(ops, message="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        
        if VerificaTotale(IFC_MATCH_CSV_OUTPUT):
            pass
            """ PER ORA RIMOSSO IN ATTESA DI SPECIFICHE
                IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT = []
                IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Nome Verifica","Stato"])
                IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Verifica Informativa - Assegnazione IFC Export As.",1])
                parameter_csv_path = os.path.join(folder, "11_XX_ValorizzazioneIFCSaveAs_Data.csv")
                with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerows(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT)
            """
        else:
            parameter_csv_path = os.path.join(folder, "11_VerificaCorrispondenzaIFC.csv")
            with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(IFC_MATCH_CSV_OUTPUT)