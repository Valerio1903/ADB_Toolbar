# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura delle viste presenti  """
__author__ = 'Roberto Dolfini'
__title__ = 'Check Nomenclatura Viste'

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

#t = Transaction(doc, "Verifica Nomenclatura Sistema")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir,'..','..','..','000_Raccolta CSV di controllo','12_CSV_NomenclaturaViste.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
VIEWS_NAMING_CSV_OUTPUT = []
VIEWS_NAMING_CSV_OUTPUT.append(["Tipo Vista","Nome Vista","ID Elemento","Stato"])
##############################################################

output.print_md("# Verifica Nomenclatura Viste")
output.print_md("---")
output.print_md("# SCRIPT IN ATTESA DI INPUT")


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

CategorieIFCDaStudiare = []
CodificaDaUsare = []
Disciplina = []
DizionarioDiVerifica = {}
CodificaCampi = []

# CHE NOMENCLATURE AVRANNO LE VISTE ?
""" 
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with open(parent_dir, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        CategorieIFCDaStudiare.append(row[0].split("-"))
        Disciplina.append(row[1])
        CodificaDaUsare.append(row[2])
        CodificaCampi.append(row[3])
for ListaCategorieIfc,CodiceDisciplina,Codifica,CodificaCampi in zip(CategorieIFCDaStudiare,Disciplina,CodificaDaUsare,CodificaCampi):
    for Categoria in ListaCategorieIfc:
        if Categoria not in DizionarioDiVerifica:
            DizionarioDiVerifica[Categoria] = [CodiceDisciplina,Codifica,CodificaCampi]
        else:
            DizionarioDiVerifica[Categoria].append([CodiceDisciplina,Codifica,CodificaCampi])
"""
"""
# COLLECTOR VISTE DI PROGETTO
All_Collector_VisteDiProgetto = FilteredElementCollector(doc).OfClass(View).ToElements()
Collector_VisteDiProgetto = []
# ESCLUSIONE PROJECT BROWSER E SYSTEM BROWSER DALLA LISTA
for x in All_Collector_VisteDiProgetto:
    if x.ViewType == ViewType.ProjectBrowser or x.ViewType == ViewType.SystemBrowser:
        pass
    else:   
        Collector_VisteDiProgetto.append(x)

output.print_md("# Verifica Nomenclature Viste")
output.print_md("---")

# VERIFICA NOMENCLATURA
DataTable = []

for Vista in Collector_VisteDiProgetto:
    if Vista:
        DataTable.append([Vista.ViewType,Vista.Name,output.linkify((Vista.Id)),":white_heavy_check_mark:"])
        VIEWS_NAMING_CSV_OUTPUT.append([Vista.ViewType,Vista.Name,Vista.Id,1])
    else:
        DataTable.append([Vista.ViewType,Vista.Name,output.linkify((Vista.Id)),":cross_mark:"])
        VIEWS_NAMING_CSV_OUTPUT.append([Vista.ViewType,Vista.Name,Vista.Id,0])

output.print_table(table_data = DataTable,title = "Verifica Nomenclatura Viste", columns = ["Tipo Vista","Nome Vista","ID Elemento","Stato"],formats = ["","","",""])
"""


"""
	#output.print_md("**ATTENZIONE, QUESTI LIVELLI NON STANNO COPY-MONITORANDO NULLA**")
	output.print_table(table_data = Levels_Status,title = "Verifica Copy-Monitor Livelli", columns = ["Id Elemento","Nome Livello","Copy-Monitor","Pin Attivo"],formats = ["","",""])
else:
	output.print_md(":cross_mark: **ATTENZIONE, NON SONO PRESENTI LIVELLI NEL MODELLO**")

"""



"""
###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(VIEWS_NAMING_CSV_OUTPUT):
            VIEWS_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            VIEWS_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Viste.",1)
        else:
            gridsandlevels_csv_path = os.path.join(folder, "12_XX_CoordinationReport_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(VIEWS_NAMING_CSV_OUTPUT)
"""
