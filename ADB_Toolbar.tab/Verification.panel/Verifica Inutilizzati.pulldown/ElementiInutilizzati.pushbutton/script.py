""" Verifica la presenza di elementi inutilizzati """

__title__ = 'Controllo Elementi\nInutilizzati'
import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
import clr
import codecs
import System
from pyrevit import forms, script
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument               
app   = __revit__.Application      
##############################################################


# CREAZIONE DELLA VISTA DI OUTPUT
output = pyrevit.output.get_output()
output.print_md("# Verifica Elementi Inutilizzati")
output.print_md("---")

# LISTA VUOTA PER CONTROLLARE TUTTE LE CATEGORIE
categories_to_check = HashSet[ElementId]()
Categorie_Inutilizzate = {}


Check = doc.GetUnusedElements(categories_to_check)
for item in Check:
    Current = doc.GetElement(item)
    try:
        Category = Current.Category.Name
        if Category not in Categorie_Inutilizzate:
            Categorie_Inutilizzate[Category] = []
        if Category != "Materiali":
            try:
                Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Famiglia":Current.Name,"Tipo":Current.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString()})
                
            except:
                Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Famiglia":Current.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsValueString(),"Tipo":Current.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString()})
                
        else:
            Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Famiglia":Current.Name,"Tipo":""})
            
    except:
        Category = "Altri Stili"
        if Category not in Categorie_Inutilizzate:
            Categorie_Inutilizzate[Category] = []

        if Current.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).HasValue:
            Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Famiglia":Current.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString(),"Tipo":Current.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString()})
            
        else:
            Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Famiglia":Current.Name,"Tipo":Current.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString()})
            
"""
# VISUALIZZAZIONE A SCHERMO
Counter = 0
for category,elements in Categorie_Inutilizzate.items():
    output.print_md("**Categoria: {0}**".format(category))
    for element in elements:
        if category != "Materiali":
            print("Famiglia: {0}, Tipo: {1},Id Elemento: {2}".format(element["Nome_Famiglia"],element["Tipo"],element["Id_Elemento"]))
            Counter +=1 
        else:
            print("Nome Materiale: {0}, Id Elemento: {1}".format(element["Nome_Famiglia"],element["Id_Elemento"]))
            Counter +=1 
"""

# GENERAZIONE TABELLA
for category, elements in Categorie_Inutilizzate.items():
    table_data = []
    for element in elements:
        row = [element["Nome_Famiglia"], output.linkify(element["Id_Elemento"]), element["Tipo"]]
        table_data.append(row)
    try:
        output.print_table(table_data=table_data, title=category, columns=["Nome Famiglia / Nome Elemento", "ID Elemento", "Tipo"])
    except Exception as e:
        print("Errore rilevato: ".format(e))

# CONTEGGIO ELEMENTI
Counter = 0
for category, elements in Categorie_Inutilizzate.items():
    Counter = Counter + len(elements)
output.print_md("---")
if Counter != 0:
    output.print_md("_Conteggio elementi inutilizzati {0}_".format(Counter))
else:
    output.print_md(":white_heavy_check_mark: **NON SONO PRESENTI ELEMENTI NON UTILIZZATI**")
output.print_md("---")



# OPZIONI ESPORTAZIONE

# GENERAZIONE TABELLA
UNUSED_ELEMENTS_CSV_DATA = []
UNUSED_ELEMENTS_CSV_DATA = [["Categoria","ID Elemento","Nome","Tipo"]]

for category, elements in Categorie_Inutilizzate.items():
    for element in elements:
        row = [category, element["Id_Elemento"],element["Nome_Famiglia"], element["Tipo"]]
        UNUSED_ELEMENTS_CSV_DATA.append(row)

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        unusedelements_csv_path = os.path.join(folder, "UnusedElements_Data.csv")
        with codecs.open(unusedelements_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(UNUSED_ELEMENTS_CSV_DATA)