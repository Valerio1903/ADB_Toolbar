# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura delle famiglie di sistema  """

__author__ = 'Roberto Dolfini'
__title__ = 'Check Nomenclatura Sistemi MEP'


# PEZZO DI CODICE UTILE PER IL FUTURO

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

t = Transaction(doc, "Verifica Nomenclatura Sistemi MEP")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir,'..','..','..','000_Raccolta CSV di controllo','12_CSV_NomenclaturaSistemiMEP.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output() 

# CREAZIONE LISTE DI OUTPUT DATA
MEP_SYSTEM_NAMING_CSV_OUTPUT = []
MEP_SYSTEM_NAMING_CSV_OUTPUT.append(["Categoria","Nome Sistema","ID Elemento","Verifica","Stato"])
##############################################################

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

def string_in_list(string, list):
    for i in list:
        if string in i:
            return True
    return False
"""
def string_in_list_contain(string, list,necessary):
    for i in list:
        if string in i:
            if necessary in i.split(string+"*")[1]:
                return True
        else:
            return False
"""

Discipline,Classe,Sottoclasse,SubdisciplineDiscipline,Raggruppamento = [],[],[],[],[]
DizionarioVerifica = {}

# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with open(parent_dir, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        RigaEsplosa = row[0].split(",")
        Discipline.append(RigaEsplosa[0])
        Classe.append(RigaEsplosa[1])
        Raggruppamento.append(RigaEsplosa[2:])
for Discipline,Classe,Raggruppamento in zip(Discipline,Classe,Raggruppamento):
    if Discipline not in DizionarioVerifica.keys():
        DizionarioVerifica[Discipline] = {}
        if Classe not in DizionarioVerifica[Discipline].keys():
            DizionarioVerifica[Discipline][Classe] = Raggruppamento
        else:
            DizionarioVerifica[Discipline][Classe].append(Raggruppamento)
    else:
        if Classe not in DizionarioVerifica[Discipline].keys():
            DizionarioVerifica[Discipline][Classe] = Raggruppamento
        else:
            DizionarioVerifica[Discipline][Classe].append(Raggruppamento)


# FILTRO PER CAPIRE QUALI HANNO UN VALORE
Filtro_TipoSistema_Pipe = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)))
Filtro_TipoSistema_Duct = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)))
Filtro_TipoSistema_CableTray = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.RBS_CTC_SERVICE_TYPE)))

Collector_PlumbingElements = FilteredElementCollector(doc).WherePasses(Filtro_TipoSistema_Pipe).WhereElementIsNotElementType().ToElements()
Collector_DuctElements = FilteredElementCollector(doc).WherePasses(Filtro_TipoSistema_Duct).WhereElementIsNotElementType().ToElements()
Collector_CableTrayElements = FilteredElementCollector(doc).WherePasses(Filtro_TipoSistema_CableTray).WhereElementIsNotElementType().ToElements()

# ESCLUSIONE DELLE CENTERLINE DAI COLLECTOR
Regroup_ElementsToVerify = []

for Elemento in Collector_PlumbingElements:
    if "CenterLine" not in  Elemento.Category.BuiltInCategory.ToString():
        Regroup_ElementsToVerify.append([Elemento.Id,Elemento.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()])

for Elemento in Collector_DuctElements:
    if "CenterLine" not in  Elemento.Category.BuiltInCategory.ToString():
        Regroup_ElementsToVerify.append([Elemento.Id,Elemento.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()])

for Elemento in Collector_CableTrayElements:
    if "CenterLine" not in  Elemento.Category.BuiltInCategory.ToString():
        Regroup_ElementsToVerify.append([Elemento.Id,Elemento.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsValueString()])

# DIZIONARIO DI VERIFICA CORRISPODENZA CODICE CLASSE CON LA CATEGORIA EFFETTIVA DELL'ELEMENTO E CORRISPONDENZA SYSTEM CLASSIFICATION
MappaturaCategorie = {
    "OST_PipeCurves": "TUB",
    "OST_PipeFitting": "TUB",
    "OST_FlexPipeCurves": "TUB",
    "OST_DuctCurves": "CND",
    "OST_DuctFitting": "CND",
    "OST_FlexDuctCurves": "CND",
    "OST_CableTray": "PAS",
    "OST_CableTrayFitting": "PAS",
}

MappaturaClassification = {
    "ID": ["Altro","Other","Domestic Cold Water","Acqua fredda sanitaria","Domestic Hot Water","Acqua calda sanitaria","Sanitary","Acque reflue"],
    "HV": ["Altro","Other","Supply Air","Aria di mandata","Return Air","Aria di ritorno","Exhaust Air","Aria di scarico","Hydronic Supply","Mandata di sistema idronico","Hydronic Return","Ritorno di sistema idronico","Vent","Ventilazione"],
    "PA": ["Altro","Other","Fire Protection Wet","Protezione antincendio a umido","Fire Protection Dry","Protezione antincendio a secco","Fire Protection Pre-Action","Protezione antincendio preattiva","Fire Protection Other","Altra protezione antincendio"]
}

# PER GLI ELEMENTI VALORIZZATI VERIFICO LA NOMENCLATURA
DataTable = []
Coppia_NomenclaturaCorretta = []
for Coppia in Regroup_ElementsToVerify:
    CategoriaElemento = doc.GetElement(Coppia[0]).Category.BuiltInCategory.ToString()
    VERIFICA = "Nomenclatura conforme."
    SIMBOLO = ":white_heavy_check_mark:"
    VALUE = 1
    if "_" not in Coppia[1]:
        VERIFICA = "Nomenclatura non conforme."
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    else:
        SuddivisioneNome = Coppia[1].split("_")
        if len(Coppia[1].split("_")) != 5:
            VERIFICA = "Lunghezza nome sistema non conforme."
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[0] not in DizionarioVerifica.keys():
            VERIFICA = "Codice disciplina errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif MappaturaCategorie[CategoriaElemento] != SuddivisioneNome[1]:
            VERIFICA = "Codice classe non conforme alla categoria dell'elemento."
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[1] not in DizionarioVerifica[SuddivisioneNome[0]].keys():
            VERIFICA = "Codice classe errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif not string_in_list(SuddivisioneNome[2], DizionarioVerifica[SuddivisioneNome[0]][SuddivisioneNome[1]]):
            VERIFICA = "Codice sottoclasse errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif string_in_list(SuddivisioneNome[2], DizionarioVerifica[SuddivisioneNome[0]][SuddivisioneNome[1]]):
            for sublist in DizionarioVerifica[SuddivisioneNome[0]][SuddivisioneNome[1]]:
                if SuddivisioneNome[2] in sublist and SuddivisioneNome[3] not in sublist:
                    VERIFICA = "Codice sottodisciplina errato"
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
                    break
                elif len(SuddivisioneNome[-1]) > 25 or " " in SuddivisioneNome[-1]:
                    VERIFICA = "Descrizione errata, verificare lunghezza e CamelCase."
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0
            if VALUE == 1 :
                Coppia_NomenclaturaCorretta.append(Coppia)
    if VALUE != 1:
        DataTable.append([doc.GetElement(Coppia[0]).Category.Name,Coppia[1],Coppia[0],VERIFICA,SIMBOLO])
        MEP_SYSTEM_NAMING_CSV_OUTPUT.append([doc.GetElement(Coppia[0]).Category.Name,Coppia[1],Coppia[0],VERIFICA,VALUE])


# SUPERATA LA VERIFICA DI NOMENCLATURA VERIFICO LA COERENZA DEL SYSTEM CLASSIFICATION 
for Coppia in Coppia_NomenclaturaCorretta:
    ClassificazioneSistema = doc.GetElement(Coppia[0]).get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsValueString()
    ChiaveRicerca = Coppia[1].split("_")[0]
    if ClassificazioneSistema not in MappaturaClassification[ChiaveRicerca]:
        VERIFICA = "Classificazione sistema errata."
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    else:
        VERIFICA = "Classificazione sistema e nomenclatura corretta."
        SIMBOLO = ":white_heavy_check_mark:"
        VALUE = 1
    #MODEL_ELEMENTS_NAMING_CSV_OUTPUT.append([Elemento.Category.Name,Elemento.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),Elemento.Id,VERIFICA,VALUE])
    DataTable.append([doc.GetElement(Coppia[0]).Category.Name,Coppia[1],Coppia[0],VERIFICA,SIMBOLO])
    MEP_SYSTEM_NAMING_CSV_OUTPUT.append([doc.GetElement(Coppia[0]).Category.Name,Coppia[1],Coppia[0],VERIFICA,VALUE])

output.print_md("# Verifica Nomenclatura Sistemi MEP")
output.print_md("---")
output.freeze()
if DataTable:
    output.print_table(table_data = DataTable, columns = ["Categoria","Nome Sistema","ID Elemento","Verifica","Stato"], formats = ["","","","",""])
else:
    output.print_md("Nessun elemento da verificare.")
output.unfreeze()


###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(MEP_SYSTEM_NAMING_CSV_OUTPUT):
            """ IN ATTESA DI INFO
            MEP_SYSTEM_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            MEP_SYSTEM_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Materiali.",1)
            """
        else:
            csv_path = os.path.join(folder, "12_Nomenclatura_SistemiMEP.csv")
            with codecs.open(csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(MEP_SYSTEM_NAMING_CSV_OUTPUT)
