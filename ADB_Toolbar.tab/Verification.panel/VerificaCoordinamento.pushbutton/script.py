""" Verifica il corretto coordinamento all'interno del modello, verificando Coordinate Condivise e Copy Monitor  """

__title__ = 'Verifica\nCoordinamento'
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

def ConvertiUnita(valore):
	return UnitUtils.Convert(valore,UnitTypeId.Feet,UnitTypeId.Meters)

######################################################################################

# VERIFICA DEL COPYMONITOR DI GRIGLIE E LIVELLI

Host_Grids = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
Host_Levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
Levels_Status,Grids_Status = [], []

COORDINATES_CSV_DATA = []
COPYMONITOR_CSV_DATA =[]

# CREAZIONE DELLA VISTA DI OUTPUT
output = pyrevit.output.get_output()


# VERIFICA DEL PIN, COORDINATE CONDIVISE, VALORI DI COORDINATE ED ANGOLO NORD REALE

LinkedFiles = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
LinkedFilesNames = [x.Name for x in LinkedFiles]

SFL = forms.SelectFromList.show(LinkedFilesNames,button_name = "Seleziona Link URS")

DocLinkSelezionato = LinkedFiles[LinkedFilesNames.index(SFL)].GetLinkDocument()
LinkSelezionato = LinkedFiles[LinkedFilesNames.index(SFL)]
Link_ProjectLocations  = DocLinkSelezionato.ProjectLocations

###
output.print_md("# Verifica Coordinamento File URS")
output.print_md("---")
output.print_md("## Coordinate Acquisite URS")
ErrorCounter = 0
###

if LinkSelezionato.Pinned:
	output.print_md(":white_heavy_check_mark: **MODELLO URS PINNATO**")
	
else:
	output.print_md(":cross_mark: **ATTENZIONE, MODELLO URS NON PINNATO !**")
	ErrorCounter += 1

if "Non condiviso" in str(SFL):
	output.print_md(":cross_mark: **ATTENZIONE, FILE A COORDINATE NON CONDIVISE!**")
	ErrorCounter += 1
else:
	output.print_md(":white_heavy_check_mark: **MODELLO URS IMPORTATO A COORDINATE CONDIVISE**")

# ESTRAGGO INFO DAL MODELLO LINKATO
Link_Iterator = Link_ProjectLocations.ForwardIterator()
l = []
while Link_Iterator.MoveNext():
    l.append(Link_Iterator.Current)
#

Link_ProjectPosition = l[0].GetProjectPosition(XYZ.Zero)

Linked_Angle = Link_ProjectPosition.Angle
Linked_EastWest = ConvertiUnita(Link_ProjectPosition.EastWest)
Linked_NorthSouth = ConvertiUnita(Link_ProjectPosition.NorthSouth)
Linked_Elevazione = ConvertiUnita(Link_ProjectPosition.Elevation)

# LA PRIMA RIGA SARA IL MODELLO LINKATO LA SECONDA IL MODELLO HOST
COORDINATES_CSV_DATA.append(["Angolo Nord Reale","Coordinata Est/Ovest","Coordinata Nord/Sud", "Elevazione"])
COORDINATES_CSV_DATA.append([Linked_Angle,Linked_EastWest,Linked_NorthSouth,Linked_Elevazione])

# ESTRAGGO INFO DAL MODELLO HOST

Host_ProjectPosition = doc.ProjectLocations
Host_Iterator = Host_ProjectPosition.ForwardIterator()

l = []
while Host_Iterator.MoveNext():
    l.append(Host_Iterator.Current)

Host_ProjectPosition = l[0].GetProjectPosition(XYZ.Zero)
Host_Angle = Host_ProjectPosition.Angle
Host_EastWest = ConvertiUnita(Host_ProjectPosition.EastWest)
Host_NorthSouth = ConvertiUnita(Host_ProjectPosition.NorthSouth)
Host_Elevazione = ConvertiUnita(Host_ProjectPosition.Elevation)

COORDINATES_CSV_DATA.append([Host_Angle,Host_EastWest,Host_NorthSouth,Host_Elevazione])

output.print_md("---")
output.print_md("## Verifica Coordinate Modelli")

if Linked_Angle == Host_Angle:
	output.print_md(":white_heavy_check_mark: **ANGOLO CONFORME TRA I DUE MODELLI**")
else:
	output.print_md(":cross_mark: **ATTENZIONE, ANGOLO NON CONFORME TRA I DUE MODELLI!**")
	output.print_md("Angolo Linkato : {0} - Angolo Host : {1}".format(Linked_Angle,Host_Angle))
	ErrorCounter += 1

if Linked_EastWest == Host_EastWest and Linked_NorthSouth == Host_NorthSouth:
	output.print_md(":white_heavy_check_mark: **COORDINATE CONFORMI TRA I DUE MODELLI**")
else:
	output.print_md(":cross_mark: **ATTENZIONE, COORDINATE NON CONFORMI TRA I DUE MODELLI!**")
	output.print_md("Coordinate Linkato E/W & N/S: {0},{1} - Coordinate Host E/W & N/S: {2},{3}".format(Linked_EastWest,Linked_NorthSouth,Host_EastWest,Host_NorthSouth))
	ErrorCounter += 1
if Linked_Elevazione == Host_Elevazione:
	output.print_md(":white_heavy_check_mark: **ELEVAZIONE CONFORME TRA I DUE MODELLI**")
else:
	output.print_md(":cross_mark: **ATTENZIONE, ELEVAZIONE NON CONFORME TRA I DUE MODELLI!**")
	output.print_md("Elevazione Linkato : {0} - Elevazione Host : {1}".format(Linked_Elevazione,Host_Elevazione))
	ErrorCounter += 1

output.print_md("---")
output.print_md("## Verifica Griglie e Livelli")

COPYMONITOR_CSV_DATA.append(["Verifica","Element ID","Nome Elemento"])
# CHECK GRIGLIE
if Host_Grids:
	for Host_Grid in Host_Grids:
		temp = []
		if Host_Grid.IsMonitoringLinkElement():
			temp.append(":white_heavy_check_mark:")
			CSV_VALUE = "Monitoring"
		else:
			temp.append(":cross_mark:")
			CSV_VALUE = "Not Monitoring"
			ErrorCounter += 1

		temp.append(output.linkify((Host_Grid.Id)))
		temp.append(Host_Grid.Name)
		#temp.append(Host_Grid.IsMonitoringLinkElement())

		COPYMONITOR_CSV_DATA.append([CSV_VALUE,Host_Grid.Id,Host_Grid.Name])
		Grids_Status.append(temp)
	#output.print_md("**ATTENZIONE, QUESTE GRIGLIE NON STANNO COPY-MONITORANDO NULLA**")
	output.print_table(table_data = Grids_Status,title = "Verifica Copy-Monitor Griglie", columns = ["Verifica","Element Id","Nome Griglia"],formats = ["","",""])
else:
	output.print_md(":cross_mark: **ATTENZIONE, NON SONO PRESENTI GRIGLIE NEL MODELLO**")

# CHECK LIVELLI
if Host_Levels:
	for Host_Level in Host_Levels:
		temp = []
		if Host_Level.IsMonitoringLinkElement():
			temp.append(":white_heavy_check_mark:")
			CSV_VALUE = "Monitoring"
		else:
			temp.append(":cross_mark:")
			CSV_VALUE = "Not Monitoring"
			ErrorCounter += 1

		temp.append(output.linkify((Host_Level.Id)))
		temp.append(Host_Level.Name)
		#temp.append(Host_Level.IsMonitoringLinkElement())

		COPYMONITOR_CSV_DATA.append([CSV_VALUE,Host_Level.Id,Host_Level.Name])
		Levels_Status.append(temp)
	#output.print_md("**ATTENZIONE, QUESTI LIVELLI NON STANNO COPY-MONITORANDO NULLA**")
	output.print_table(table_data = Levels_Status,title = "Verifica Copy-Monitor Livelli", columns = ["Verifica","Element Id","Nome Livello"],formats = ["","",""])
else:
	output.print_md(":cross_mark: **ATTENZIONE, NON SONO PRESENTI LIVELLI NEL MODELLO**")


output.print_md("---")
if ErrorCounter == 1:
	output.print_md("_E' STATO RILEVATO 1 ERRORE_!")
elif ErrorCounter > 1:
	output.print_md("_SONO PRESENTI {0} ERRORI!_".format(ErrorCounter))
else:
	output.print_md("_NON SONO STATI RILEVATI ERRORI_")
output.print_md("---")

###OPZIONI ESPORTAZIONE
ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":

	folder = pyrevit.forms.pick_folder()

	if folder:
		copymonitor_csv_path = os.path.join(folder, "CopyMonitorReport_Data.csv")
		coordination_csv_path = os.path.join(folder, "CoordinationReport_Data.csv")
		with open(copymonitor_csv_path, mode='w') as file:
			writer = csv.writer(file)
			writer.writerows(COPYMONITOR_CSV_DATA)

		with open(coordination_csv_path, mode='w') as file:
			writer = csv.writer(file)
			writer.writerows(COORDINATES_CSV_DATA)
