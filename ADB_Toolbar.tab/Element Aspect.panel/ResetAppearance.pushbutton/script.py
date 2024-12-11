""" Resetta ogni modifica all'aspetto degli elementi in vista.  """

__title__ = 'Resetta\nAspetto'

import codecs
import re
import unicodedata
#import pyrevit
#from pyrevit import *
import math
import clr
import System
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
aview = doc.ActiveView

t = Transaction(doc, "Reset Appearance")

##############################################################

ElementiVisibili = FilteredElementCollector(doc,aview.Id).WhereElementIsNotElementType().ToElements()
Override_CLEAR = OverrideGraphicSettings()

t.Start()

for Elemento in ElementiVisibili:
	aview.SetElementOverrides(Elemento.Id,Override_CLEAR)

t.Commit()
