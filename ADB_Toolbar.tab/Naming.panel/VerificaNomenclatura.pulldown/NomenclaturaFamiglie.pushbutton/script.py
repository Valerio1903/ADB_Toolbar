# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura delle famiglie caricabili """
__title__ = 'Check Nomenclatura Caricabili'


# PEZZO DI CODICE UTILE PER IL FUTURO
"""
def EBICENOME(Chiave):
    category_mapping = {}
    doc   = __revit__.ActiveUIDocument.Document
    for bic in Enum.GetValues(BuiltInCategory):
        try:
            category = doc.Settings.Categories.get_Item(bic)
            
            if category and "XML" not in str(bic):
                if category.Name not in category_mapping:
                    localized_name = category.Name
                    if localized_name and not any(substring in localized_name for substring in ["<", "XML"]):
                        category_mapping[category.Name] = bic
                else:
                    if localized_name and not any(substring in localized_name for substring in ["<", "XML"]):
                        category_mapping[category.Name].append(bic)
                    pass
        except Exception as e:
            #category_mapping[bic] = "Error: {}".format(e)
            pass
    return category_mapping[Chiave]


    # COSTRUISCO LA LISTA DELLE MIE CATEGORIE DI RIFERIMENTO
    ListaCategorie = []
    for bic, localized_name in category_mapping.items():
        if localized_name and not any(substring in localized_name for substring in ["<", "XML"]) and "XML" not in str(bic):
            ListaCategorie.append((bic, localized_name))
    ListaCategorie.sort(key=lambda x: x[1])
    return ListaCategorie
"""
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
import csv
import codecs
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
ActiveView = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Nomenclatura Sistema")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir,'..', '..','..','000_Raccolta CSV di controllo','12_CSV_NomenclaturaFamiglie.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
MODEL_ELEMENTS_NAMING_CSV_OUTPUT = []
DataTable = []
MODEL_ELEMENTS_NAMING_CSV_OUTPUT.append(["Categoria","Nome Famiglia","Nome Tipo","ID Elemento","Verifica Nome Famiglia","Verifica Nome Tipo","Stato"])
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

def build_regex(pattern):
    """
    Converts a pattern string into a regular expression.
    'n' represents a digit, other characters are treated as literals.
    """
    parts = []
    i = 0
    while i < len(pattern):
        if pattern[i] == "n":
            # Count consecutive 'n's for digit sequence
            count = 0
            while i < len(pattern) and pattern[i] == "n":
                count += 1
                i += 1
            parts.append(r"\d{" + str(count) + "}")
        else:
            # Treat non-'n' characters as literals
            literal = ""
            while i < len(pattern) and pattern[i] != "n":
                literal += pattern[i]
                i += 1
            parts.append(re.escape(literal))  # Escape special regex characters
    return "".join(parts)

def check_value(rule, value):
    """
    Checks if the value matches any pattern in the rule string.
    Patterns are separated by '/'.
    """
    alternatives = rule.split("/")
    return any(re.match(build_regex(alt) + "$", value) for alt in alternatives)

ActiveView = doc.ActiveView.Id
Collector = FilteredElementCollector(doc,ActiveView).WhereElementIsNotElementType().ToElements()

#! GENERAZIONE DEL DIZIONARIO DI VERIFICA
def process_column_6(value):
    if "-" in value:
        fields = value.split("-")
        return [field.split("~") for field in fields]
    return value.split("~")

file_path = r"C:\Users\2Dto6D\Desktop\12_CSV_Nomenclatura Famiglie.csv"

# Read the CSV file
try:
    with open(file_path, 'r') as file:
        lines = file.readlines()
except IOError:
    print("Error: File '{}' not found on desktop.".format(file_path))
    exit(1)

# Process the first row into MATERIALI list
MATERIALI = [code.strip() for code in lines[0].strip().split(",")]
DISCIPLINA = [code.strip() for code in lines[1].strip().split(",")]
# Initialize the naming system list
NAMING_SYSTEM = []

for line in lines[2:]:
    fields = [field.strip() for field in line.strip().split("@")]
    num_fields = len(fields)
    
    if num_fields not in [6, 7, 8]:
        print("Warning: Row '{}' has {} fields, expected 6, 7, or 8. Skipping.".format(line.strip(), num_fields))
        continue
    
    min_max = fields[1].split("/")
    if len(min_max) != 2:
        print("Warning: Min/Max field '{}' in row '{}' is invalid. Skipping.".format(fields[1], line.strip()))
        continue
    try:
        min_value = int(min_max[0])
        max_value = int(min_max[1])

    except ValueError:
        print("Warning: Min/Max values '{}' in '{}' are not integers. Skipping.".format(min_max[0], min_max[1], line.strip()))
        continue
    
    classe_elemento = fields[2]
    materiali = MATERIALI if fields[3] == "TRUE" else []
    codes = process_column_6(fields[-2])
    built_in_categories = fields[-1]
    
    # Handle Number and Format fields
    if fields[4] == "NULL":
        number = []  # Use empty list for no pattern
        try:
            formato = int(fields[5]) if fields[5] != "NULL" else None
        except ValueError:
            formato = None
    else:
        number = fields[4]  # e.g., "nnnn[mm]"
        try:
            formato = int(fields[5]) if fields[5] != "NULL" else None
        except ValueError:
            formato = None
    if fields[6]== "NULL":
        codes = []
    entry = {
        "Disciplina": DISCIPLINA,
        "Min": min_value,
        "Max": max_value,
        "Classe": classe_elemento,  
        "Materiali": materiali,
        "Format": formato,  
        "Number": number, 
        "SottoClasse/Subdisciplina": codes,
        "BuiltInCategories": built_in_categories
    }
    NAMING_SYSTEM.append(entry)

#! INIZIO VERIFICA NOME FAMIGLIA

for Elemento in Collector:
    Scelta = None
    Dizionario_Scelto = None  # Initialize to avoid undefined variable
    if Elemento.Category:
        for x in Elemento.GetDependentElements(None):
            if "Panel" in doc.GetElement(x).GetType().ToString():
                Scelta = "OST_Roof-Panel"
        for entry in NAMING_SYSTEM:
            if not Scelta and Elemento.Category.BuiltInCategory.ToString() in entry["BuiltInCategories"]:
                Dizionario_Scelto = entry
                break  # Exit loop once matched
            elif Scelta and Scelta in entry["BuiltInCategories"]:
                Dizionario_Scelto = entry
                break
    if not Dizionario_Scelto:
        #if Elemento.Category:
            #print("non ho trovato questo :{}".format(Elemento.Category.BuiltInCategory.ToString()))
        continue  

    #? INIZIO DEFINIZIONE 

    Nome_Famiglia = EstraiInfoOggetto(Elemento)[0]
    Nome_Tipo = EstraiInfoOggetto(Elemento)[1]

    if Nome_Famiglia and isinstance(Nome_Famiglia, str):
        SPLIT_Nome_Famiglia = Nome_Famiglia.split("_")
    else:
        SPLIT_Nome_Famiglia = False

    SPLIT_Nome_Tipo = Nome_Tipo.split("_")

    VERIFICA = "Nomenclatura corretta"
    SIMBOLO = ":white_heavy_check_mark:"
    VALUE = 1
    VERIFICA_TIPO_CARICABILE = ""
    SIMBOLO_TIPO_CARICABILE = ""
    VALUE_TIPO_CARICABILE = None

    ValoreVerifica = SPLIT_Nome_Tipo
    if isinstance(SPLIT_Nome_Famiglia, list):  #! CON QUESTA ECCEZIONE PRENDO IL NOME DELLE FAMIGLIE
        ValoreVerifica = SPLIT_Nome_Famiglia

    #! Inizio Verifica
    if len(ValoreVerifica) < Dizionario_Scelto["Min"]:
        VERIFICA = "Nome famiglia troppo corto"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif len(ValoreVerifica) > Dizionario_Scelto["Max"]:
        VERIFICA = "Nome famiglia troppo lungo"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif ValoreVerifica[0] not in Dizionario_Scelto["Disciplina"]:
        VERIFICA = "Codice disciplina non conforme"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif ValoreVerifica[1] != Dizionario_Scelto["Classe"]:
        VERIFICA = "Codice classe non conforme"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif ValoreVerifica[2] not in Dizionario_Scelto["SottoClasse/Subdisciplina"]:
        if not Elemento.Category.BuiltInCategory.ToString() in ["OST_CableTray","OST_PipeCurves"]:  
            VERIFICA = "Codice sottoclasse o subdisciplina non conforme"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif Dizionario_Scelto["Materiali"] and ValoreVerifica[2] not in Dizionario_Scelto["Materiali"]:
            VERIFICA = "Codice materiale non conforme"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
    elif Dizionario_Scelto["Materiali"] and ValoreVerifica[3] not in Dizionario_Scelto["Materiali"]:
        VERIFICA = "Codice materiale non conforme"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif Elemento.Category.BuiltInCategory.ToString() in ["OST_Walls","OST_Floors","OST_Ceilings","OST_Roofs"] and not Dizionario_Scelto["Materiali"] and Dizionario_Scelto["Number"] != [] and not check_value(Dizionario_Scelto["Number"], ValoreVerifica[3]):
        VERIFICA = "Codice dimensionale non conforme"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    elif Dizionario_Scelto["Materiali"]:
        try:
            if ValoreVerifica[4] in Dizionario_Scelto["Materiali"] and Dizionario_Scelto["Number"] != [] and not check_value(Dizionario_Scelto["Number"], ValoreVerifica[4]):
                VERIFICA = "Codice dimensionale non conforme"
                SIMBOLO = ":cross_mark:"
                VALUE = 0
        except:
            if Dizionario_Scelto["Materiali"] and ValoreVerifica[3] in Dizionario_Scelto["Materiali"] and Dizionario_Scelto["Number"] != [] and not check_value(Dizionario_Scelto["Number"], ValoreVerifica[3]):
                VERIFICA = "Codice dimensionale non conforme"
                SIMBOLO = ":cross_mark:"
                VALUE = 0
    
    #! ECCEZIONI SPECIFICHE PER CATEGORIE SPECIFICHE

    #! PORTE E FINESTE
    elif Elemento.Category.BuiltInCategory.ToString() in ["OST_Doors","OST_Windows"]:
        if len(ValoreVerifica) >= 5:
            if len(ValoreVerifica[-1]) != 2:
                if len(ValoreVerifica[-1]) > 25:
                    VERIFICA = "Descrizione troppo lunga"
                    SIMBOLO = ":cross_mark:"
                    VALUE = 0  
            elif not check_value("nA",ValoreVerifica[-1]):
                VERIFICA = "Verificare il numero ante"
                SIMBOLO = ":cross_mark:"
                VALUE = 0
    elif len(ValoreVerifica[-1]) > Dizionario_Scelto["Format"]:
        VERIFICA = "Descrizione troppo lunga"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    Nome_Famiglia = Elemento.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()

    #! QUI INIZIA LA VERIFICA DELLA NOMENCLATURA DEI TIPI 
    if VALUE == 1 and isinstance(Elemento,FamilyInstance):
        VERIFICA_TIPO_CARICABILE = "Nomenclatura corretta"
        SIMBOLO_TIPO_CARICABILE = ":white_heavy_check_mark:"
        VALUE_TIPO_CARICABILE = 1

        if Nome_Famiglia not in Nome_Tipo:
            VERIFICA_TIPO_CARICABILE = "Nome famiglia non presente nel tipo"
            SIMBOLO_TIPO_CARICABILE = ":cross_mark:"
            VALUE_TIPO_CARICABILE = 0
        else:
            SPLITFAMIGLIA_Nome_Tipo = Nome_Tipo.split(Nome_Famiglia)[1].split("_")
            try:
                if check_value("nnnn/nnnnxnnnn/nnnxnnn/nnnxnnnn/nnnnxnnn",SPLITFAMIGLIA_Nome_Tipo[1]) == False:
                    VERIFICA_TIPO_CARICABILE = "Dimensione non conforme"
                    SIMBOLO_TIPO_CARICABILE = ":cross_mark:"
                    VALUE_TIPO_CARICABILE = 0
                elif len(SPLITFAMIGLIA_Nome_Tipo[-1]) > 25:
                    VERIFICA_TIPO_CARICABILE = "Descrizione troppo lunga"
                    SIMBOLO_TIPO_CARICABILE = ":cross_mark:"
                    VALUE_TIPO_CARICABILE = 0
            except:
                    VERIFICA_TIPO_CARICABILE = "Dimensione non presente"
                    SIMBOLO_TIPO_CARICABILE = ":cross_mark:"
                    VALUE_TIPO_CARICABILE = 0
    if any([VALUE == 0, VALUE_TIPO_CARICABILE == 0]):
        VALUE = 0
    else:
        VALUE = 1
    MODEL_ELEMENTS_NAMING_CSV_OUTPUT.append([
        Elemento.Category.Name if Elemento.Category else "N/A",
        Nome_Famiglia,
        Nome_Tipo,
        Elemento.Id.ToString(),
        VERIFICA,
        VERIFICA_TIPO_CARICABILE,
        VALUE
    ])

    DataTable.append([
        Elemento.Category.Name if Elemento.Category else "N/A",
        Nome_Famiglia,
        Nome_Tipo,
        output.linkify(Elemento.Id),
        VERIFICA,
        SIMBOLO,
        VERIFICA_TIPO_CARICABILE,
        SIMBOLO_TIPO_CARICABILE
    ])

output.print_table(table_data = DataTable,title = "Verifica Nomenclatura Famiglie", columns = ["Categoria","Nome Famiglia","Nome Tipo","ID Elemento","Verifica","Stato","Verifica Tipo(Caricabili)","Stato Tipo(Caricabili)"])







###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(MODEL_ELEMENTS_NAMING_CSV_OUTPUT):
            pass
            """ IN ATTESA DI CONFERME 
            GRID_LEVELS_NAMING_CSV_OUTPUT = []
            GRID_LEVELS_NAMING_CSV_OUTPUT.append(["Nome Verifica","Stato"])
            GRID_LEVELS_NAMING_CSV_OUTPUT.append(["Naming Convention - Nomenclatura Griglie e Livelli.",1])
            gridsandlevels_csv_path = os.path.join(folder, "12_GriglieELivelli_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(GRID_LEVELS_NAMING_CSV_OUTPUT)
            """
        else:
            familynaming_csv_path = os.path.join(folder, "12_NomenclaturaFamiglie_Data.csv")
            with codecs.open(familynaming_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(MODEL_ELEMENTS_NAMING_CSV_OUTPUT)











