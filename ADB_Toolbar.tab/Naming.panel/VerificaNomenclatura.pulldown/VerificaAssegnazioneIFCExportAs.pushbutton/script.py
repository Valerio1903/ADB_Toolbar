# -*- coding: utf-8 -*-

""" Verifica la valorizzazione del parametro ExportIFCAs  """

__title__ = 'Check Valorizzazione IFC Export As'

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

t = Transaction(doc, "Colorazione IFCSaveAs")
##############################################################

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()
output.print_md("# Verifica Valorizzazione IFC Export As")
output.print_md("---")
# CREAZIONE LISTE DI OUTPUT DATA
IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT = []
IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Categoria","Nome Famiglia","Nome Tipo","ID Elemento","Stato"])
##############################################################

def EstraiInfoOggetto(oggetto):
    # SE L'OGGETTO E' UNA FAMIGLIA CARICABILE MI PRENDE FAMIGLIA E TIPO
    if isinstance(oggetto,FamilyInstance):
        return [oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
    # SE L'OGGETTO E' UNA FAMIGLIA DI SISTEMA MI PRENDE SOLO IL TIPO
    else:
        try:
            #return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            return [oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
        except:
            try:
                return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
            except:
                return "NON TROVATO"

# DEFINIZIONE COLORI PER TROUBLESHOOTING
rosso = Color(255,100, 100)
arancio = Color(255, 180, 100)
viola = Color(180,100,255)
bianco = Color(255,255,255)
SolidPattern = FilteredElementCollector(doc).OfClass(FillPatternElement).FirstElement()

Override_Sistema = OverrideGraphicSettings()
#Override_Sistema.SetProjectionLineColor(bianco)
Override_Sistema.SetSurfaceForegroundPatternId(SolidPattern.Id)
Override_Sistema.SetSurfaceForegroundPatternColor(rosso)
#
Override_Caricabile = OverrideGraphicSettings()
#Override_Caricabile.SetProjectionLineColor(bianco)
Override_Caricabile.SetSurfaceForegroundPatternId(SolidPattern.Id)
Override_Caricabile.SetSurfaceForegroundPatternColor(arancio)
#
Override_Locale = OverrideGraphicSettings()
#Override_Locale.SetProjectionLineColor(bianco)
Override_Locale.SetSurfaceForegroundPatternId(SolidPattern.Id)
Override_Locale.SetSurfaceForegroundPatternColor(viola)
#
Override_Corretto = OverrideGraphicSettings()
Override_Corretto.SetHalftone(True)



# FILTRO ELEMENTI VALORIZZATI
Filtro_AssegnazioneEffettuata = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)))
Collector_IFC_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_AssegnazioneEffettuata).WhereElementIsNotElementType().ToElements()

# FILTRO PER CAPIRE QUALI NON SONO STATI VALORIZZATI
Filtro_MancataAssegnazione = ElementParameterFilter(FilterStringRule(ParameterValueProvider(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)), FilterStringEquals(), ""))
Temp_Collector_IFC_NON_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_MancataAssegnazione).WhereElementIsNotElementType().ToElements()
Collector_IFC_NON_Valorizzato = [x for x in Temp_Collector_IFC_NON_Valorizzato if x.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS) != None]

Lista_Elementi_IFC_NON_Valorizzato = []
for Elemento in Collector_IFC_NON_Valorizzato:
    if Elemento :
        try:
            if Elemento.Category.CategoryType == CategoryType.Model:
                if Elemento.Category is not None:
                    if Elemento.Category.HasMaterialQuantities or Elemento.Category.Name in ['Air Terminals', 'Ducts', 'Duct Fittings', 'Duct Accessories', 'Duct Insulations', 'Duct Linings', "Flex Ducts", 'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Pipe Insulations', "Flex Pipes", 'Cable Trays', 'Cable Tray Fittings', 'Conduits', 'Conduit Fittings']:
                        if "CenterLine" not in str(Elemento.Category.BuiltInCategory):
                            Lista_Elementi_IFC_NON_Valorizzato.append(Elemento)
        except Exception as e:
            print(Cat)
            print(e)
            pass


DataTable = []

t.Start()

for Elemento in Lista_Elementi_IFC_NON_Valorizzato:
    # ESTRAGGO INFORMAZIONI
    Categoria = Elemento.Category.Name
    ID_Elemento = Elemento.Id
    Nome_Famiglia = EstraiInfoOggetto(Elemento)[0]
    Nome_Tipo = EstraiInfoOggetto(Elemento)[1]
    Tipologia = "Sistema"
    aview.SetElementOverrides(ID_Elemento, Override_Sistema)
    if isinstance(Elemento,FamilyInstance):
        if Elemento.Symbol.Family.IsInPlace:
            Tipologia = "Locale"
            aview.SetElementOverrides(ID_Elemento, Override_Locale)
        else:
            Tipologia = "Caricabile"
            aview.SetElementOverrides(ID_Elemento, Override_Caricabile)
    
    DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,ID_Elemento,Tipologia])
    IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,0])

for Corretto in Collector_IFC_Valorizzato:
    try:
        aview.SetElementOverrides(Corretto.Id, Override_Corretto)
    except Exception as e:
        print(e)
        pass

t.Commit()

for Elemento in Collector_IFC_Valorizzato:
    # ESTRAGGO INFORMAZIONI
    Categoria = Elemento.Category.Name
    ID_Elemento = Elemento.Id
    Nome_Famiglia = EstraiInfoOggetto(Elemento)[0]
    Nome_Tipo = EstraiInfoOggetto(Elemento)[1]
    IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,1])

# CREAZIONE DELLA VISTA DI OUTPUT

OrderedData = sorted(DataTable, key=lambda status: status[2])
if len(OrderedData) != 0:
    output.freeze()
    output = pyrevit.output.get_output()
    output.print_md("# Verifica Elementi 'IFCSaveAs' Non Valorizzati")
    output.print_md("---")
    output.print_table(table_data = OrderedData, columns = ["Nome Famiglia", "Nome Tipo","Categoria", "ID Elemento","Tipologia"], formats = ["","","","",""])
    output.print_md("---")
    output.unfreeze()
else:
    output.print_md(":white_heavy_check_mark: **TUTTI GLI ELEMENTI SONO VALORIZZATI CORRETTAMENTE**")
###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si", "No"]
Scelta = pyrevit.forms.CommandSwitchWindow.show(ops, message="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT):
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT = []
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Nome Verifica","Stato"])
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Verifica Informativa - Assegnazione IFC Export As.",1])
            parameter_csv_path = os.path.join(folder, "11_XX_ValorizzazioneIFCSaveAs_Data.csv")
            with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT)
        else:
            parameter_csv_path = os.path.join(folder, "11_XX_ValorizzazioneIFCSaveAs_Data.csv")
            with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT)