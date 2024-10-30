""" Check di verifica dei materiali inutilizzati """

__title__ = 'Unused Materials\nCheck'
import codecs
import re
import unicodedata
from pyrevit import forms
import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

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

# SCRIPT START : 

#categories_to_check = HashSet[ElementId]()
categories_to_check = HashSet[ElementId]()
categories_to_check.Add(ElementId(BuiltInCategory.OST_Materials))

Check = doc.GetUnusedElements(categories_to_check)

print([doc.GetElement(x).Name for x in Check])