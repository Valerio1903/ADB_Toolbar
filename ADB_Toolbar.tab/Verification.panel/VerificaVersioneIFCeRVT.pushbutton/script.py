# -*- coding: utf-8 -*-

""" Verifica la versione dei files IFC e RVT selezionati.  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Versione Files'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import forms, script
import math
import clr
import os
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import codecs
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
#PER IL FLOATING POINT
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument               
app   = __revit__.Application      
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Versione Files")

##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
#script_dir = os.path.dirname(__file__)
#parent_dir = os.path.abspath(os.path.join(script_dir, '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))


# CREAZIONE LISTE DI OUTPUT DATA
VERIFICAVERSIONE_CSV_OUTPUT = []
VERIFICAVERSIONE_CSV_OUTPUT.append(["Nome File","Versione File","Verifica", "Stato"])

output = pyrevit.output.get_output()

Cartella_Controllo = pyrevit.forms.pick_folder()

DataTable = []

if Cartella_Controllo:
    for NomeFile in os.listdir(Cartella_Controllo):
        PercorsoFile = os.path.join(Cartella_Controllo, NomeFile)

        if ".ifc" in NomeFile:
            with open(PercorsoFile, 'rb') as file:
                
                for riga in file:
                    if "FILE_NAME" in riga:
                        Nome = riga.split("'")[1].split("'")[0]
                    elif "FILE_SCHEMA" in riga:
                        Versione = riga.split("'")[1].split("'")[0]
                        if "4x3" not in Versione.lower():
                            DataTable.append([Nome, Versione,"Versione non conforme, usare IFC 4x3" ,":cross_mark:"])
                            VERIFICAVERSIONE_CSV_OUTPUT.append([Nome, Versione,"Versione non conforme - usare IFC 4x3",0])
                        else:
                            DataTable.append([Nome, Versione,"Versione conforme.",":white_heavy_check_mark:"])
                            VERIFICAVERSIONE_CSV_OUTPUT.append([Nome, Versione,"Versione conforme.",1])
        elif ".rvt" in NomeFile:
            Versione = BasicFileInfo.Extract(PercorsoFile).Format
            Nome = NomeFile
            if "2024" not in Versione:
                DataTable.append([Nome, Versione,"Versione non conforme, usare 2024" , ":cross_mark:"])
                VERIFICAVERSIONE_CSV_OUTPUT.append([Nome, Versione,"Versione non conforme - usare 20245",0])
            else:
                DataTable.append([Nome, Versione,"Versione conforme.", ":white_heavy_check_mark:"])
                VERIFICAVERSIONE_CSV_OUTPUT.append([Nome, Versione,"Versione conforme.",1])


output.print_md("# Verifica Versione Files")
output.print_md("---")
if len(DataTable) > 0:
    output.freeze()
    output.print_table(table_data = DataTable,columns = ["Nome File","Versione File","Verifica","Stato"],formats = ["","","",""])
    output.unfreeze()
else:
    output.print_md("**Nessun file IFC o RVT trovato nella cartella selezionata.**")
    script.exit()

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        copymonitor_csv_path = os.path.join(folder, "10_VersioneFile_Data.csv")
        with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(VERIFICAVERSIONE_CSV_OUTPUT)
        if VerificaTotale(VERIFICAVERSIONE_CSV_OUTPUT):
            pass
            """
            VERIFICAUNITA_CSV_OUTPUT = []
            VERIFICAUNITA_CSV_OUTPUT.append(["Nome Verifica","Stato"])
            VERIFICAUNITA_CSV_OUTPUT.append(["Georeferenziazione e Orientamento - CopyMonitor correttamente effettuato.",1])
            copymonitor_csv_path = os.path.join(folder, "07_02_CopyMonitorReport_Data.csv")
            with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(VERIFICAUNITA_CSV_OUTPUT)
            """
        else:
            pass
