# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura delle famiglie di sistema  """
__author__ = 'Roberto Dolfini'
__title__ = 'Check Nomenclatura Sistema'


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

t = Transaction(doc, "Verifica Nomenclatura Sistema")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
#script_dir = os.path.dirname(__file__)
#parent_dir = os.path.abspath(os.path.join(script_dir,'..','..','..','000_Raccolta CSV di controllo','12_CSV_NomenclaturaSistema.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
MODEL_ELEMENTS_NAMING_CSV_OUTPUT = []
MODEL_ELEMENTS_NAMING_CSV_OUTPUT.append(["Categoria","Nome Tipo","ID Elemento","Verifica","Stato"])
##############################################################

output.print_md("# Verifica Nomenclatura Famiglie di Sistema")
output.print_md("---")
output.print_md("# SCRIPT IN ATTESA DI INPUT")


"""
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

# FILTRO PER CAPIRE QUALI HANNO UN VALORE
Filtro_AssegnazioneEffettuata = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)))
Collector_IFC_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_AssegnazioneEffettuata).WhereElementIsNotElementType().ToElements()

# FILTRO PER CAPIRE QUALI NON SONO STATI VALORIZZATI
Filtro_MancataAssegnazione = ElementParameterFilter(FilterStringRule(ParameterValueProvider(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)), FilterStringEquals(), ""))
Temp_Collector_IFC_NON_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_MancataAssegnazione).WhereElementIsNotElementType().ToElements()
Collector_IFC_NON_Valorizzato = [x for x in Temp_Collector_IFC_NON_Valorizzato if x.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS) != None]

# PER GLI ELEMENTI VALORIZZATI VERIFICO LA NOMENCLATURA
for Elemento in Collector_IFC_Valorizzato:
    if not EstraiInfoOggetto(Elemento)[0]:
            
        Valore_IFCExportAs = Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsValueString()
        # PRIMA VERIFICO LA PRESENZA IN DATABASE
        if Valore_IFCExportAs not in DizionarioDiVerifica.keys():
            print("Chiave {} non presente nel database.".format(Valore_IFCExportAs))
            print("---")
        else:
            # ESTRAGGO LE INFORMAZIONI DALL'ELEMENTO
            Scomposizione_Nome_Famiglia = EstraiInfoOggetto(Elemento)[1].split("_")
            Nome_Tipo = EstraiInfoOggetto(Elemento)[1]


            # ESTRAGGO I DATI DAL DIZIONARIO E VERIFICO IL NOME FAMIGLIA
            try:
                EstrazioneDati = DizionarioDiVerifica[Valore_IFCExportAs]
                V0_VerificaCampi = EstrazioneDati[-1]
                V1_CodiceDisciplina = EstrazioneDati[0]
                V2_CodiceClasse = EstrazioneDati[1].split("*")[0]
                V3_CodiciSottoClasse = EstrazioneDati[1].split("*")[1].split("-")
                ValoriVerifica = [V1_CodiceDisciplina,V2_CodiceClasse,V3_CodiciSottoClasse]
                # INIZIO CON LA VERIFICA
                if not VerificaCodifica(Nome_Tipo, V0_VerificaCampi)[0]:
                    print(VerificaCodifica(Nome_Tipo, V0_VerificaCampi)[1])
                    pass
                else:
                    # RAGGRUPPO LE VERIFICHE IN UNA LISTA PER CAPIRE QUALE CAMPO FALLISCE
                    CheckStatus = [Scomposizione_Nome_Famiglia[0] == V1_CodiceDisciplina, Scomposizione_Nome_Famiglia[1] == V2_CodiceClasse, Scomposizione_Nome_Famiglia[2] in V3_CodiciSottoClasse]
                    DefCheck = True
                    for index,Check in enumerate(CheckStatus):
                        if not Check:
                            DefCheck = [False,index]
                            print("Nomenclatura errata al campo {}, {} - Classe Ifc: {}".format(index+1,EstraiInfoOggetto(Elemento)[0],Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsValueString()))
                            break
                    else:
                        print("Nomenclatura Famiglia corretta {}, {} - Classe Ifc: {}".format(Elemento.Category.Name,EstraiInfoOggetto(Elemento)[0],Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsValueString()))
            except Exception as e:
                print(e)
                pass
"""