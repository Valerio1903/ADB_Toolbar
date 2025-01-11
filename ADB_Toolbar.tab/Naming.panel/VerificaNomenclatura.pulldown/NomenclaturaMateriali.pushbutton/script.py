# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura dei materiali presenti  """
__author__= 'Roberto Dolfini'
__title__ = 'Check Nomenclatura Materiali'

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

t = Transaction(doc, "Verifica Nomenclatura Materiali")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, '..','Raccolta CSV di controllo','Database_CodiciMateriali.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
MATERIAL_NAMING_CSV_OUTPUT = []
MATERIAL_NAMING_CSV_OUTPUT.append(["Nome Materiale","ID Elemento","Verifica","Stato"])
##############################################################

def VerificaCodifica(nome, riferimento):
    mappa_regex = {'X': '[A-Z]', 'x': '[a-z]', 'n': '\\d', 'A': 'A'}
    
    # DIVIDI IN BLOCCHI BASANDOTI SUL DELIMITATORE (ES. '_')
    blocchi_nome = nome.split('_')
    blocchi_riferimento = riferimento.split('_')
        
    # CONTROLLA SE IL NUMERO DI BLOCCHI CORRISPONDE
    if len(blocchi_nome) != len(blocchi_riferimento):
        return False, "Numero blocchi errato: Attesi {}, Trovati {}. Nome : {}".format(len(blocchi_riferimento), len(blocchi_nome),nome)
        
    # CONFRONTA BLOCCO PER BLOCCO
    for i, (blocco_nome, blocco_ref) in enumerate(zip(blocchi_nome, blocchi_riferimento)):
        if i == 3 and blocco_nome == "All":  # ECCEZIONE PER IL BLOCCO 3
            continue
        regex = ''.join(mappa_regex.get(char, re.escape(char)) for char in blocco_ref)
        if not re.match("^{}$".format(regex), blocco_nome):  # SIMULA FULLMATCH
            return False, "Errore nel blocco {} ('{}'): Atteso '{}'. Nome : {}".format(i + 1, blocco_nome, blocco_ref,nome)
        
    return True, "Il nome rispetta il formato."

Codice = []
Descrizione = []
DizionarioDiVerifica = {}
CodificaCampi = []
Riferimento = []


# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with open(parent_dir, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        Codice.append(row[0])
        Descrizione.append(row[1])
        Riferimento.append(row[-1])

for codice,descrizione,riferimento in zip(Codice,Descrizione,Riferimento):
    DizionarioDiVerifica[codice] = [descrizione,riferimento]
    FormatNome = riferimento
# COLLECTOR VISTE DI PROGETTO
Collector_Materiali = FilteredElementCollector(doc).OfClass(Material).ToElements()

output.print_md("# Verifica Nomenclatura Materiali")
output.print_md("---")
DefaultKey = DizionarioDiVerifica.keys()[0]

DataTable = []
Stato = 0

for Materiale in Collector_Materiali:
    # VERIFICA NUMERO CAMPI
    if len(Materiale.Name.split("_")) < 2:
        VERIFICA = "NUMERO CAMPI ERRATO"
        SIMBOLO = ":cross_mark:"
        DataTable.append([Materiale.Name, output.linkify(Materiale.Id), VERIFICA, SIMBOLO])
        MATERIAL_NAMING_CSV_OUTPUT.append([Materiale.Name, Materiale.Id, VERIFICA, Stato]) 
        continue
    
    campi = Materiale.Name.split("_")
    CodiceCliente = campi[0]
    CodiceMateriale = campi[1]
    DescrizioneMateriale = "_".join(campi[2:]) if len(campi) > 2 else ""
    
    # VERIFICA PRESENZA ADB A INIZIO NOME
    if "ADB" not in CodiceCliente:
        VERIFICA = "CODICE ADB MANCANTE"
        SIMBOLO = ":cross_mark:"
        DataTable.append([Materiale.Name, output.linkify(Materiale.Id), VERIFICA])
        MATERIAL_NAMING_CSV_OUTPUT.append([Materiale.Name, Materiale.Id, VERIFICA, Stato])
        continue

    # Verifica Codice Materiale nel dizionario
    if CodiceMateriale not in DizionarioDiVerifica:
        VERIFICA = "CODICE NON TROVATO NEL CSV"
        SIMBOLO = ":question_mark:"
        DataTable.append([Materiale.Name, output.linkify(Materiale.Id), VERIFICA, SIMBOLO])
        MATERIAL_NAMING_CSV_OUTPUT.append([Materiale.Name, Materiale.Id, VERIFICA, Stato])
        continue
    
    # Recupera i dati dal dizionario
    dati_materiale = DizionarioDiVerifica[CodiceMateriale]
    descrizione = DizionarioDiVerifica[CodiceMateriale][0]
    formato_atteso = FormatNome
    
    # VERIFICA CAMEL CASE
    codice_valido, msg = VerificaCodifica(CodiceMateriale, formato_atteso)
    if not codice_valido:
        VERIFICA = "CODICE NON IN CAMEL CASE ({msg})".format(msg=msg)
        SIMBOLO = ":cross_mark:"
        DataTable.append([Materiale.Name, output.linkify(Materiale.Id), VERIFICA, SIMBOLO])
        MATERIAL_NAMING_CSV_OUTPUT.append([Materiale.Name, Materiale.Id, VERIFICA, Stato])
        continue

    # VERIFICA DESCRIZIONE
    if DescrizioneMateriale == descrizione:
        VERIFICA = "Corretto"
        SIMBOLO = ":white_heavy_check_mark:"
        Stato = 1
    elif len(DescrizioneMateriale) > 25:
        VERIFICA = "DESCRIZIONE TROPPO LUNGA"
        SIMBOLO = ":cross_mark:"
    else:
        VERIFICA = "DESCRIZIONE NON CORRISPONDENTE"
        SIMBOLO = ":question_mark:"

    DataTable.append([Materiale.Name, output.linkify(Materiale.Id), VERIFICA,SIMBOLO])
    MATERIAL_NAMING_CSV_OUTPUT.append([Materiale.Name, Materiale.Id, VERIFICA, Stato])

output.freeze()
output.print_table(
    table_data=DataTable,
    title="Verifica Nomenclatura Materiali",
    columns=["Nome Materiale", "ID Elemento", "Verifica", "Stato"],
    formats=["", "", "", ""]
)             
output.unfreeze()


###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(MATERIAL_NAMING_CSV_OUTPUT):
            MATERIAL_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            MATERIAL_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Materiali.",1)
        else:
            csv_path = os.path.join(folder, "12_CSV_Nomenclatura_Materiali.csv")
            with codecs.open(csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(MATERIAL_NAMING_CSV_OUTPUT)
