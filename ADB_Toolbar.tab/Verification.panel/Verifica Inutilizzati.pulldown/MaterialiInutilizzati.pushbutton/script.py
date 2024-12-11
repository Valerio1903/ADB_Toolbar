""" Verifica la presenza di materiali inutilizzati """

__title__ = 'Controllo Materiali\nInutilizzati'
import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
import clr
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
output.print_md("# Verifica Materiali Inutilizzati")
output.print_md("---")

# LISTA VUOTA PER CONTROLLARE TUTTE LE CATEGORIE
categories_to_check = HashSet[ElementId]()
categories_to_check.Add(ElementId(BuiltInCategory.OST_Materials))
Categorie_Inutilizzate = {}
Counter = 0
UNUSED_MATERIALS_CSV_DATA = []
UNUSED_MATERIALS_CSV_DATA = [["Categoria","ID Elemento","Nome"]]
Check = doc.GetUnusedElements(categories_to_check)
for item in Check:
    
    Current = doc.GetElement(item)
    Category = Current.Category.Name
    if Category not in Categorie_Inutilizzate:
        Categorie_Inutilizzate[Category] = []
    Categorie_Inutilizzate[Category].append({"Id_Elemento":item,"Nome_Materiale":Current.Name})
    Counter +=1

    # ARRICCHIMENTO FILE CSV
    UNUSED_MATERIALS_CSV_DATA.append([Current.Category.Name,item,Current.Name])

"""
# VISUALIZZAZIONE A SCHERMO DEFAULT 
Counter = 0
for category,elements in Categorie_Inutilizzate.items():
    output.print_md("**Categoria: {0}**".format(category))
    for element in elements:
        print("Nome Materiale: {0}, Id Elemento: {1}".format(element["Nome_Materiale"],element["Id_Elemento"]))
        Counter +=1 
"""

# GENERAZIONE TABELLA
for category, elements in Categorie_Inutilizzate.items():
    table_data = []
    for element in elements:
        row = [element["Nome_Materiale"], output.linkify(element["Id_Elemento"])]
        table_data.append(row)
    try:
        output.print_table(table_data=table_data, columns=["Nome Materiale", "ID Elemento"])
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

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    
    if folder:
        if VerificaTotale(UNUSED_MATERIALS_CSV_DATA):
            UNUSED_MATERIALS_CSV_DATA.append("Nome Verifica","Stato")
            UNUSED_MATERIALS_CSV_DATA.append("Integrità e pulizia file - Non sono presenti materiali inutilizzati.",1)
        else:
            unusedmaterials_csv_path = os.path.join(folder, "UnusedMaterials_Data.csv")
            with codecs.open(unusedmaterials_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(UNUSED_MATERIALS_CSV_DATA)