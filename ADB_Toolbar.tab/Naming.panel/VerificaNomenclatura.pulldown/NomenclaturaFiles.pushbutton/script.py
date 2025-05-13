# -*- coding: utf-8 -*-

""" Verifica la nomenclatura e che la dimensione dei files presenti in directory non superi i 200MB  """
__author__ = 'Roberto Dolfini'
__title__ = 'Check Nomenclatura e Dimensione Files'



######################################
import Autodesk.Revit
import Autodesk.Revit.DB
import Autodesk.Revit.DB.Architecture
import pyrevit
from pyrevit import forms, script
import codecs
import re
from re import *
from Autodesk.Revit.DB import BuiltInCategory, Category, CategoryType
import System
from System import Enum
from System.Collections.Generic import Dictionary
import unicodedata
import math
import clr
import os
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import time
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
#t = Transaction(doc, "TESTING")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir,'..', '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Files.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
FILEDIMENSION_CSV_OUTPUT = []
FILEDIMENSION_CSV_OUTPUT.append(["Nome File","Verifica","Dimensione","Stato"])
##############################################################


def format_size(size_in_bytes):
    # Define the units for file sizes
    units = ['B', 'KB', 'MB', 'GB', 'TB']
    # Calculate the index for the appropriate unit
    unit_index = 0
    size = size_in_bytes
    while size >= 1024 and unit_index < len(units) - 1:
        size /= 1024.0
        unit_index += 1
    # Return the size formatted to 3 decimal places along with the unit
    return "{:.3f} {}".format(float(size), units[unit_index])


DizionarioDiVerifica = {}
DizionarioDiVerifica["Zona"]=[]
DizionarioDiVerifica["Fase"]=[]
DizionarioDiVerifica["Sistema/Argomento"]=[]
DizionarioDiVerifica["Tipo"]=[]
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    rows = list(reader)  # Convert to a list of rows
    if len(rows) >= 2:  # Ensure CSV has at least 2 rows
        # Correctly split codes by commas for each row
        DizionarioDiVerifica["Zona"] = rows[0]
        DizionarioDiVerifica["Fase"] = rows[1]
        DizionarioDiVerifica["Sistema/Argomento"] = rows[2]
        DizionarioDiVerifica["Tipo"] = rows[3]
    else:
        output.print_md(":warning: Errore con il file CSV di controllo!")


DataTable = []

Cartella_Controllo = pyrevit.forms.pick_folder()
if Cartella_Controllo:
    DimensioneMax = 200 * 1024 * 1024
    
    for NomeFile in os.listdir(Cartella_Controllo):
        PercorsoFile = os.path.join(Cartella_Controllo, NomeFile)

        if os.path.isfile(PercorsoFile):
            DimensioneFile = os.path.getsize(PercorsoFile)

            if len(NomeFile.split("_")) < 8:
                    VERIFICA = "Lunghezza nome file non conforme."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
            else:
                VERIFICA = "Nomenclatura corretta."
                SIMBOLO = ":white_heavy_check_mark:"
                VALUE = 1
                VERIFICA_DIMENSIONE = ":cross_mark:"
                SuddivisioneNome = NomeFile.split("_")
                if "9nnn" not in SuddivisioneNome[0]:
                    VERIFICA = "Codice progressivo errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif "J.002" not in SuddivisioneNome[1]:
                    VERIFICA = "WBS di progetto errata."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not any(SuddivisioneNome[2] in zona for zona in DizionarioDiVerifica["Zona"]):
                    VERIFICA = "Codice zona errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not any(SuddivisioneNome[3] in fase for fase in DizionarioDiVerifica["Fase"]):
                    VERIFICA = "Codice fase errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not any(SuddivisioneNome[4] in sistema for sistema in DizionarioDiVerifica["Sistema/Argomento"]):
                    VERIFICA = "Codice sistema/argomento errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not any(SuddivisioneNome[5] in tipo for tipo in DizionarioDiVerifica["Tipo"]):
                    VERIFICA = "Codice tipo errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not(SuddivisioneNome[6].isnumeric()) and len(SuddivisioneNome[6]) != 2 or len(SuddivisioneNome[6]) != 2:
                    VERIFICA = "Codice progressivo errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                elif not(SuddivisioneNome[7].split(".")[0].isnumeric()) and len(SuddivisioneNome[7]) != 2 or len(SuddivisioneNome[7].split(".")[0]) != 2:
                    VERIFICA = "Codice revisione errato."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                # CONTROLLO DELLE DIMENSIONI FILE
            if DimensioneFile < DimensioneMax:
                VERIFICA_DIMENSIONE = ":white_heavy_check_mark:"
                #DataTable.append([NomeFile,format_size(DimensioneFile),":white_heavy_check_mark:"])
            else:
                VERIFICA_DIMENSIONE = ":cross_mark:"
                #DataTable.append([NomeFile,format_size(DimensioneFile),":cross_mark:"])
            # CONTROLLO NOME DEL FILE
            DataTable.append([NomeFile,format_size(DimensioneFile),VERIFICA_DIMENSIONE,VERIFICA,SIMBOLO])
            FILEDIMENSION_CSV_OUTPUT.append([NomeFile,VERIFICA,format_size(DimensioneFile),VALUE])

### ASSEGNARE LA NOMENCLATURA CORRETTA DEI FILESx\\

output = pyrevit.output.get_output()
output.print_md("# Verifica Dimensione Files")
output.print_md("---")

output.print_table(table_data = DataTable,columns = ["Nome File","Dimensione File","< 200 MB","Verifica Nomenclatura","Esito Verifica"],formats = ["","","","",""])

###OPZIONI ESPORTAZIONE
ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        parameter_csv_path = os.path.join(folder, "12_DimensioneNomenclaturaFile_Data.csv")
    with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(FILEDIMENSION_CSV_OUTPUT)