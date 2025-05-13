# -*- coding: utf-8 -*-

""" Verifica che gli elementi siano correttamente vincolati tra livello e livello.  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Interruzione elementi tra livelli.'

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

#t = Transaction(doc, "Verifica Versione Files")

##############################################################

# CREAZIONE LISTE DI OUTPUT DATA
VINCOLOLIVELLI_CSV_OUTPUT = []
VINCOLOLIVELLI_CSV_OUTPUT.append(["Categoria","Id Elemento","Livello Base","Livello Superiore","Offset Base","Offset Superiore","Verifica","Stato"])


# Funzione: recupera livelli ordinati per elevazione e suddivisi per disciplina
def recupera_livelli(doc):
    livelli = FilteredElementCollector(doc).OfClass(Level).ToElements()
    livelli_con_altitudine = [(liv.Id, liv.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble()) for liv in livelli]
    livelli_ordinati = sorted(livelli_con_altitudine, key=lambda x: x[1])

    livelli_arc, livelli_str, livelli_ignoti = [], [], []
    for livello_id, _ in livelli_ordinati:
        livello = doc.GetElement(livello_id)
        codice = livello.Name.split("_")[0]
        if "AR" in codice:
            livelli_arc.append(livello)
        elif "ST" in codice:
            livelli_str.append(livello)
        else:
            livelli_ignoti.append(livello)

    return livelli_arc, livelli_str, livelli_ignoti

#ESTRAZIONE MURI - COLONNE ED INFO
CatFilter = List[BuiltInCategory]()
CatFilter.Add(BuiltInCategory.OST_Walls)
CatFilter.Add(BuiltInCategory.OST_Columns)
CatFilter.Add(BuiltInCategory.OST_StructuralColumns)

FiltroCategorie = ElementMulticategoryFilter(CatFilter)

ElementiVerifica = FilteredElementCollector(doc).WherePasses(FiltroCategorie).WhereElementIsNotElementType().ToElements()

DataTable = []

livelli_arc, livelli_str, livelli_ignoti = recupera_livelli(doc)

for Elemento in ElementiVerifica:
    try:
        Base = Elemento.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        Superiore = Elemento.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()
        NomeBase = Elemento.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsValueString()
        NomeSuperiore = Elemento.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsValueString()
        base_offset = Elemento.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        param_offset = Elemento.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        bot_offset_value = base_offset.AsDouble() if base_offset and base_offset.HasValue else 0.0
        top_offset_value = param_offset.AsDouble() if param_offset and param_offset.HasValue else 0.0
        OffsetBase = round(bot_offset_value*0.3048, 3) if bot_offset_value else 0.0
        OffsetSuperiore = round(top_offset_value*0.3048, 3) if top_offset_value else 0.0

    except:
        Base = Elemento.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()
        Superiore = Elemento.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()
        NomeBase = Elemento.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString()
        NomeSuperiore = Elemento.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsValueString() 
        base_offset = Elemento.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        param_offset = Elemento.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
        bot_offset_value = base_offset.AsDouble() if base_offset and base_offset.HasValue else 0.0
        top_offset_value = param_offset.AsDouble() if param_offset and param_offset.HasValue else 0.0
        OffsetBase = round(bot_offset_value*0.3048, 3) if bot_offset_value else 0.0
        OffsetSuperiore = round(top_offset_value*0.3048, 3) if top_offset_value else 0.0

    VALUE = 1
    VERIFICA = "Elemento vincolato correttamente tra livelli."
    SIMBOLO = ":white_heavy_check_mark:"

    if "AR" in NomeBase and "AR" in NomeSuperiore:
        IDLivelliAR = [livello.Id for livello in livelli_arc]
        if Superiore != IDLivelliAR[IDLivelliAR.index(Base)+1]:
            VALUE = 0
            VERIFICA = "Elemento non vincolato correttamente tra livelli."
            SIMBOLO = ":cross_mark:"
        
    elif "ST" in NomeBase and "ST" in NomeSuperiore:
        IDLivelliST = [livello.Id for livello in livelli_str]
        if Superiore != IDLivelliST[IDLivelliST.index(Base)+1]:
            VALUE = 0
            VERIFICA = "Elemento non vincolato correttamente tra livelli."
            SIMBOLO = ":cross_mark:"
        
    elif "ST" in NomeBase and "AR" in NomeSuperiore or "AR" in NomeBase and "ST" in NomeSuperiore:
        VALUE = 0
        VERIFICA = "Elemento vincolato tra livelli di discipline diverse."
        SIMBOLO = ":cross_mark:"
        
    elif Superiore == ElementId.InvalidElementId:
        VALUE = 0
        VERIFICA = "Elemento senza vincolo superiore."
        SIMBOLO = ":cross_mark:"

    
    DataTable.append([Elemento.Category.Name,output.linkify(Elemento.Id),NomeBase,NomeSuperiore,OffsetBase,OffsetSuperiore,VERIFICA,SIMBOLO])
    VINCOLOLIVELLI_CSV_OUTPUT.append([Elemento.Category.Name,Elemento.Id,NomeBase,NomeSuperiore,OffsetBase,OffsetSuperiore,VERIFICA,VALUE])

output.freeze()
output.print_md("# Verifica vincolo elementi verticali")
output.print_md("---")
output.print_table(table_data = DataTable,columns = ["Categoria","Id Elemento","Livello Base","Livello Superiore","Offset Superiore","Verifica","Stato"],formats = ["","","","","",""])
output.unfreeze()


ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        copymonitor_csv_path = os.path.join(folder, "13_VincoloTraLivelli_Data.csv")
        with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(VINCOLOLIVELLI_CSV_OUTPUT)
