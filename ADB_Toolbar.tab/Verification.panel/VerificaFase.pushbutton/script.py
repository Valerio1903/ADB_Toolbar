""" Verifica l'assegnazione della fase corretta ai diversi elementi presenti in vista  """

__title__ = 'Verifica Fase\nElementi'

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
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument               
app   = __revit__.Application      
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Fase")

##############################################################

# CREAZIONE LISTE DI OUTPUT DATA
VERIFICAFASE_CSV_OUTPUT = []
VERIFICAFASE_CSV_OUTPUT.append(["Nome Elemento", "ID Elemento","Categoria","Fase", "Stato"])

# INPUT DA PARTE DELL'UTENTE DEI VALORI DI VERIFICA
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,Separator, Button, CheckBox)
components = [
            Label('Nome fase da ricercare :'),
            TextBox('fase_nome', Text="Stato di Fatto"),
            Button('Select'),
            ]
form = FlexForm('Definisci la fase da ricercare', components)
form.show()

if not form.values:
    script.exit()
      
ValoriUtente = form.values
FaseScelta = ValoriUtente['fase_nome']

def EstraiFamigliaOggetto(oggetto):
    try:
        return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    except:
        try:
            return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
        except:
            return "NON TROVATO"
        
output = pyrevit.output.get_output()
output.print_md("# Verifica Fase Elementi")
output.print_md("---")

#VERIFICO LA PRESENZA DELLA FASE ALL'INTERNO DEL FILE

FasiDelDocumento = FilteredElementCollector(doc).OfClass(Phase)
FasiPresenti = [Fase.Name for Fase in FasiDelDocumento]

if FaseScelta not in FasiPresenti:
    output.print_md(":cross_mark: **La fase {} non si trova all'interno del file.** :cross_mark:".format(FaseScelta))
    VERIFICAFASE_CSV_OUTPUT.append(["LA FASE {} NON SI TROVA ALL'INTERNO DEL FILE".format(FaseScelta)]) # CAPIRE COME ESPORTARLA IN MODO EFFICACE
    script.exit()

ElementiInVista = FilteredElementCollector(doc, aview.Id).WhereElementIsNotElementType().ToElements()
DataTable = []
    
for Elemento in ElementiInVista:
    if Elemento.HasPhases():
        try:
            Fase = Elemento.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()
            if FaseScelta in Fase:
                VERIFICAFASE_CSV_OUTPUT.append([EstraiFamigliaOggetto(Elemento),Elemento.Id,Elemento.Category.Name,Fase,1])
            else:
                DataTable.append([EstraiFamigliaOggetto(Elemento),output.linkify(Elemento.Id),Elemento.Category.Name,Fase,"Fase non corrispondente"])
                VERIFICAFASE_CSV_OUTPUT.append([EstraiFamigliaOggetto(Elemento),Elemento.Id,Elemento.Category.Name,Fase,0])
        except Exception as e:
            print(e)
            pass


# CREAZIONE DELLA VISTA DI OUTPUT
if len(DataTable) != 0:
    output.freeze()
    OrderedData = sorted(DataTable, key=lambda status: status[2])
    output.print_table(table_data = OrderedData, columns = ["Nome Elemento", "ID Elemento","Categoria","Fase", "Stato"], formats = ["","","","",""])
    output.unfreeze()
else:
    output.print_md(":white_heavy_check_mark: **Tutti gli elementi sono nella fase corretta.** :white_heavy_check_mark:")
###OPZIONI ESPORTAZIONE
ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        parameter_csv_path = os.path.join(folder, "FaseErrata_Data.csv")
    with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(VERIFICAFASE_CSV_OUTPUT)




###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    
    if folder:
        if VerificaTotale(VERIFICAFASE_CSV_OUTPUT):
            VERIFICAFASE_CSV_OUTPUT.append("Nome Verifica","Stato")
            VERIFICAFASE_CSV_OUTPUT.append("Verifica Informativa - Gli elementi sono valorizzati correttamente.",1)
        else:
            unusedmaterials_csv_path = os.path.join(folder, "11_01_FaseErrata_Data.csv")
            with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(VERIFICAFASE_CSV_OUTPUT)





















"""
else:
    ElementiInVista = FilteredElementCollector(doc, aview.Id).WhereElementIsNotElementType().ToElements()
    FaseElementi, DataTable = [], []
    
    for Elemento in ElementiInVista:
        IDFase = Elemento.get_Parameter(BuiltInParameter.PHASE_CREATED).AsElementId()
        Fase = doc.GetElement(IDFase)
        if Fase:
            FaseElementi.append((Elemento.Id, Fase.Name))   

    for element_id, Fase_name in FaseElementi:
        ElementSingolo = doc.GetElement(element_id)
        if Fase_name not in FaseScelta:
            DataTable.append([EstraiFamigliaOggetto(ElementSingolo),output.linkify(element_id),ElementSingolo.Category.Name,Fase_name,"Non corrisponde",0])
            VERIFICAFASE_CSV_OUTPUT.append([EstraiFamigliaOggetto(ElementSingolo),element_id,ElementSingolo.Category.Name,Fase_name,0])
            
        else:
            DataTable.append([EstraiFamigliaOggetto(ElementSingolo),output.linkify(element_id),ElementSingolo.Category.Name,Fase_name,"Corrispondenza",1])
            VERIFICAFASE_CSV_OUTPUT.append([EstraiFamigliaOggetto(ElementSingolo),element_id,ElementSingolo.Category.Name,Fase_name,1])


    # CREAZIONE DELLA VISTA DI OUTPUT
    OrderedData = sorted(DataTable, key=lambda status: status[-1])
    output.print_table(table_data = OrderedData, columns = ["Nome Elemento", "ID Elemento","Categoria","Fase", "Stato"], formats = ["","","","",""])
    
    ###OPZIONI ESPORTAZIONE
    ops = ["Si","No"]
    Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
    if Scelta == "Si":
        folder = pyrevit.forms.pick_folder()
        if folder:
            parameter_csv_path = os.path.join(folder, "FaseErrata_Data.csv")
        with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(VERIFICAFASE_CSV_OUTPUT)

"""
