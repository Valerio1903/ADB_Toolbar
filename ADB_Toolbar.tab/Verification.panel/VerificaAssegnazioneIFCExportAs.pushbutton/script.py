# -*- coding: utf-8 -*-

""" Verifica la valorizzazione del parametro ExportIFCAs  """

__author__ = 'Roberto Dolfini'
__title__ = 'Check Valorizzazione\nIFC Export As'

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

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
csv_dir = os.path.join(script_dir, 'IfcClass_PredfType.csv')


#parent_dir = os.path.abspath(os.path.join(script_dir, '..','Raccolta CSV di controllo','Database_StrutturaNomenclaturaCaricabile.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()
# CREAZIONE LISTE DI OUTPUT DATA
IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT = []
IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append(["Categoria","Nome Famiglia","Nome Tipo","ID Elemento","Verifica","Stato"])
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

def VerificaCorrispondenza(Categoria, Current_IfcClass, Current_Predef, DizionarioDiVerifica):

    result = [
        "<b style='color:red;'>IFC Class Errata :{}</b><br>".format(Current_IfcClass),
        "<b style='color:red;'>IFC Predef Errato :{}</b><br>".format(Current_Predef)
    ]

    if Categoria in DizionarioDiVerifica:
        for sublist in DizionarioDiVerifica[Categoria]:
            if Current_IfcClass.strip().lower() in map(str.lower, sublist):
                result[0] = ":white_heavy_check_mark:"
            if Current_Predef.strip().lower() in map(str.lower, sublist):
                result[1] = ":white_heavy_check_mark:"
            if result[0] == ":white_heavy_check_mark:" and result[1] == ":white_heavy_check_mark:":
                break

    return result




RVT_Cat, RVT_Built, IFCCLASS, IFCPREDEFINEDTYPE, Descrizione = [],[],[],[],[]
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
DizionarioDiVerifica = {}
with open(csv_dir, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        if len(row) == 5:
            RVT_Cat.append(row[0])
            RVT_Built.append(row[1])
            IFCCLASS.append(row[2])
            IFCPREDEFINEDTYPE.append(row[3])
            Descrizione.append(row[4])
    
        
for Categoria,IfcClass,PredType,DescrizioneCampi in zip(RVT_Built,IFCCLASS,IFCPREDEFINEDTYPE,Descrizione):
    if Categoria not in DizionarioDiVerifica:
        DizionarioDiVerifica[Categoria] = []
        DizionarioDiVerifica[Categoria].append([IfcClass,PredType,DescrizioneCampi])
    else:
        DizionarioDiVerifica[Categoria].append([IfcClass,PredType,DescrizioneCampi])

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
Temp_Collector_IFC_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_AssegnazioneEffettuata).WhereElementIsNotElementType().ToElements()
Collector_IFC_Valorizzato = []
# VERIFICA INCROCIATA PER EVITARE VALORI VUOTI 
for elemento in Temp_Collector_IFC_Valorizzato:
    if elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsString() != "":
        Collector_IFC_Valorizzato.append(elemento)
        

# FILTRO PER CAPIRE QUALI NON SONO STATI VALORIZZATI
Filtro_MancataAssegnazione = ElementParameterFilter(FilterStringRule(ParameterValueProvider(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)), FilterStringEquals(), ""))
Temp_Collector_IFC_NON_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_MancataAssegnazione).WhereElementIsNotElementType().ToElements()
Collector_IFC_NON_Valorizzato = [x for x in Temp_Collector_IFC_NON_Valorizzato if x.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS) != None]

MEPBuiltInCategories = [
    "OST_ElectricalFixtures",
    "OST_LightingFixtures",
    "OST_LightingDevices",
    "OST_CableTray",
    "OST_CableTrayFitting",
    "OST_Conduit",
    "OST_ConduitFitting",
    "OST_DataDevices",
    "OST_FireAlarmDevices",
    "OST_NurseCallDevices",
    "OST_SecurityDevices",
    "OST_TelephoneDevices",
    "OST_CommunicationDevices",
    "OST_DuctCurves",
    "OST_FlexDuctCurves",
    "OST_DuctFitting",
    "OST_DuctAccessory",
    "OST_DuctInsulations",
    "OST_DuctLinings",
    "OST_MechanicalEquipment",
    "OST_MechanicalControlDevice",
    "OST_PipeCurves",
    "OST_FlexPipeCurves",
    "OST_PipeFitting",
    "OST_PipeAccessory",
    "OST_PlumbingFixtures",
    "OST_PlumbingEquipment",
    "OST_PipeInsulations",
    "OST_Sprinklers",
    "OST_DuctTerminal",
    "OST_ElectricalEquipment",
]

Lista_Elementi_IFC_NON_Valorizzato = []
for Elemento in Collector_IFC_NON_Valorizzato:
    if Elemento :
        try:
            if Elemento.Category.CategoryType == CategoryType.Model and Elemento.Category is not None :
                if Elemento.Category.HasMaterialQuantities or str(Elemento.Category.BuiltInCategory) in MEPBuiltInCategories:
                    if "CenterLine" not in str(Elemento.Category.BuiltInCategory):
                        Lista_Elementi_IFC_NON_Valorizzato.append(Elemento)
                elif "Rail" in str(Elemento):
                    Lista_Elementi_IFC_NON_Valorizzato.append(Elemento)
                else:
                    pass
        except Exception as e:
            print(Elemento.Category.CategoryType)
            print(e)
            pass
    else:
        print("Non elemento".format(Elemento.Id))

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
    else:
        pass
    Verifica = "IFC Class non valorizzato"
    IFC_Class_Verifica = "<b style=color:orange;'>IFC Class non valorizzato</b><br>"
    IFC_Predef_Verifica = "<b style=color:orange;'>Ifc Predefined Type non presente</b><br>"
    DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])
    IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,0])

for Corretto in Collector_IFC_Valorizzato:
    try:
        aview.SetElementOverrides(Corretto.Id, Override_Corretto)
    except Exception as e:
        print(e)
        pass

t.Commit()

for Elemento in Collector_IFC_Valorizzato:
    # ESTRAGGO INFORMAZIONI
    Categoria = str(Elemento.Category.BuiltInCategory)
    ID_Elemento = Elemento.Id
    Nome_Famiglia = EstraiInfoOggetto(Elemento)[0]
    Nome_Tipo = EstraiInfoOggetto(Elemento)[1]
    Current_Predef = Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE).AsValueString()
    Current_IfcClass = Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsValueString()
    Tipologia = "Sistema"
    IFC_Class_Verifica = VerificaCorrispondenza(Categoria, Current_IfcClass, Current_Predef, DizionarioDiVerifica)[0]
    IFC_Predef_Verifica = VerificaCorrispondenza(Categoria, Current_IfcClass, Current_Predef, DizionarioDiVerifica)[1]
    if isinstance(Elemento,FamilyInstance):
        if Elemento.Symbol.Family.IsInPlace:
            Tipologia = "Locale"
        else:
            Tipologia = "Caricabile"

    # VERIFICA VALORIZZAZIONE PREDEFINED TYPE
    ## Verifica se il valore è presente

    if Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE).AsValueString() == None or Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE).AsValueString() == "":
        Verifica = "IFC PREDEF TYPE non valorizzato"
        IFC_Predef_Verifica = "<b style=color:red;'>Ifc Predefined Type non presente</b><br>"
        DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])
        IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,0])
    else:
        
        if IFC_Class_Verifica == ":white_heavy_check_mark:" and IFC_Predef_Verifica == ":white_heavy_check_mark:":
            Verifica = "Corretto"
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,1])
            #DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])
        elif IFC_Class_Verifica == ":white_heavy_check_mark:" and IFC_Predef_Verifica == ":cross_mark:":
            Verifica = "IFC PREDEF TYPE errato"
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,0])
            DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])
        elif IFC_Class_Verifica == ":cross_mark:" and IFC_Predef_Verifica == ":white_heavy_check_mark:":
            Verifica = "IFC CLASS errata"
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,0])
            DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])
        else:
            Verifica = "IFC CLASS e IFC PREDEF TYPE errati"
            IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Verifica,0])
            DataTable.append([Nome_Famiglia,Nome_Tipo,Categoria,output.linkify(ID_Elemento),Tipologia,IFC_Class_Verifica,IFC_Predef_Verifica])

output.freeze()
output = pyrevit.output.get_output()
output.print_md("# Verifica Elementi 'IFCSaveAs'")
output.print_md("---")
output.print_table(table_data = DataTable, columns = ["Nome Famiglia", "Nome Tipo","Categoria", "ID Elemento","Tipologia","IFC Class","IFC Predef Type"])
output.resize(1500, 900)
output.unfreeze()








###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si", "No"]
Scelta = pyrevit.forms.CommandSwitchWindow.show(ops, message="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        
        if VerificaTotale(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT):
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
            parameter_csv_path = os.path.join(folder, "11_XX_ValorizzazioneIFCSaveAs_Data.csv")
            with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(IFC_EXPORT_AS_PARAMETER_CSV_OUTPUT)


