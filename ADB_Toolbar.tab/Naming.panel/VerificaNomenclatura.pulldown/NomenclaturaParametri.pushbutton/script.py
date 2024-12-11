# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura dei parametri  """
__title__ = 'Check Nomenclatura Parametri'

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

#t = Transaction(doc, "Verifica Nomenclatura Griglie e Livelli")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, '..','Raccolta CSV di controllo','Database_CodiciParametriProvvisorio.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
PARAMETER_NAMING_CSV_OUTPUT = []
PARAMETER_NAMING_CSV_OUTPUT.append(["Nome Parametro","ID Elemento","Verifica","Stato"])

##############################################################

output.print_md("# Verifica Nomenclatura Parametri")
output.print_md("---")


# TUTTE ISTANZE
# Colleziona tutti gli elementi nel documento
all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

# Filtro per HasMaterialQuantities e prendo le categorie che hanno questo metodo più altre che mi servono
model_elements = []
for i in all_elements :
    if i.Category is not None:
        if i.Category.HasMaterialQuantities or i.Category.Name in ['Air Terminals', 'Ducts', 'Duct Fittings', 'Duct Accessories', 'Duct Insulations', 'Duct Linings', "Flex Ducts", 'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Pipe Insulations', "Flex Pipes", 'Cable Trays', 'Cable Tray Fittings', 'Conduits', 'Conduit Fittings']:
            model_elements.append(i)



ID_TipiASchermo = []
TipiASchermo = []
# PRENDO ELEMENTI A SCHERMO PER ANALISI DEL TIPO
# TIPI DI ELEMENTI EFFETTIVAMENTE POSIZIONATI NEL PROGETTO

for Elemento in model_elements:
    ID_TipiASchermo.append(Elemento.GetTypeId())
Unique_ID_TipiASchermo = list(set(ID_TipiASchermo)) 

for ID in Unique_ID_TipiASchermo:
    TipiASchermo.append(doc.GetElement(ID))



DizionarioDiVerifica = {}
DizionarioDiVerifica["Pset"]=[]
DizionarioDiVerifica["ClasseParametro"]=[]
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    rows = list(reader)  # Convert to a list of rows
    if len(rows) >= 2:  # Ensure CSV has at least 2 rows
        # Correctly split codes by commas for each row
        DizionarioDiVerifica["Pset"] = rows[0]
        DizionarioDiVerifica["ClasseParametro"] = rows[1]
    else:
        output.print_md(":warning: CSV file does not have enough rows for verification!")



# ESTRAGGO I PARAMETRI DI TIPO ESCLUDENDO I BUILTIN
CSV_DataTable = []

for Tipo in TipiASchermo:
    if Tipo:
        Parametri_Tipo = Tipo.Parameters
        ID_Tipo = Tipo.Id
        Nome_Tipo = Tipo.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString()
        Iterator = Parametri_Tipo.ForwardIterator()
        while Iterator.MoveNext():
            Parametro = Iterator.Current
            # AVVIO IL CHECK DEI PARAMETRI
            if Parametro.HasValue and Parametro.Definition.Id.Value > 0:
                NomeParametro = Parametro.Definition.Name
                if len(NomeParametro.split("_")) < 4:
                    VERIFICA = "Lunghezza nome parametro non conforme."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0 
                else:
                    VERIFICA = ""
                    SIMBOLO = 1
                    VALUE = 1
                    SuddivisioneNome = NomeParametro.split("_")
                    if SuddivisioneNome[0] not in DizionarioDiVerifica["Pset"]:
                        VERIFICA = "Codice Pset errato."
                        SIMBOLO = ":cross_mark:"
                        VALUE = 0
                    elif SuddivisioneNome[1] not in DizionarioDiVerifica["ClasseParametro"]:
                        VERIFICA = "Classe parametro errata."
                        SIMBOLO = ":cross_mark:"
                        VALUE = 0
                    elif "T" not in SuddivisioneNome[2]:
                        VERIFICA = "Codice tipologia parametro errato."
                        SIMBOLO = ":cross_mark:"
                        VALUE = 0
                    elif "_" not in NomeParametro:
                        VERIFICA = "Nome parametro non conforme."
                        SIMBOLO = ":cross_mark:"
                        VALUE = 0
                    elif len(SuddivisioneNome[3]) > 25:
                        VERIFICA = "Descrizione parametro superiore ai 25 caratteri."
                        SIMBOLO = ":cross_mark:"
                        VALUE = 0
                    else:
                        pass
                    """
                        VERIFICA = "Nomenclatura parametro conforme."
                        SIMBOLO = ":white_check_mark:"
                        VALUE = 1
                    """
                if VALUE != 1:
                    CSV_DataTable.append([Nome_Tipo,ID_Tipo,NomeParametro, VERIFICA, VALUE])
    
output.freeze()
print(CSV_DataTable)
output.unfreeze()



###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Definire esportazione file CSV?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(CSV_DataTable):
            PARAMETER_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            PARAMETER_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Parametri.",1)
        else:
            csv_path = os.path.join(folder, "XX_XX_ParameterNaming_Data.csv")
            with codecs.open(csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(CSV_DataTable)




























"""
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        Chiave=row[0]
        Codifica=row[1].split("~")
        DizionarioDiVerifica[Chiave] = Codifica

VERIFICA = "NUMERO CAMPI ERRATO"
SIMBOLO = ":cross_mark:"


# Append results to DataTable for both errors and successes
DataTable.append([Livello.Category.Name, Nome, output.linkify((Livello.Id)), VERIFICA, SIMBOLO])
GRID_LEVELS_NAMING_CSV_OUTPUT.append([Livello.Category.Name, Livello.Name, Livello.Id, VALUE])

output.print_table(table_data = DataTable,title = "Verifica Nomenclatura Livelli", columns = ["Categoria","ID Elemento","Verifica"],formats = ["","",""])

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(GRID_LEVELS_NAMING_CSV_OUTPUT):
            GRID_LEVELS_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            GRID_LEVELS_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Griglie e Livelli.",1)
        else:
            gridsandlevels_csv_path = os.path.join(folder, "12_XX_CoordinationReport_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(GRID_LEVELS_NAMING_CSV_OUTPUT)

"""