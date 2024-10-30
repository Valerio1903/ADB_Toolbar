"""Riallinea le tubazioni e crea un raccordo"""


__title__ = 'Adjust and create\n fitting'
__author__ = 'Valerio Mascia'

# -*- coding: utf-8 -*-
import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Ottieni il documento attivo da PyRevit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Funzione per trovare il punto piu lontano
def FarPoint(point, line):
    Extremes = [line.GetEndPoint(0), line.GetEndPoint(1)]
    Distances = [point.DistanceTo(Extremes[0]), point.DistanceTo(Extremes[1])]
    return Extremes[Distances.index(max(Distances))]

# Funzione per trovare la distanza piu corta e il punto piu vicino tra due curve
def ClosestPointAndDDistance(curve1, curve2):
    start1 = curve1.GetEndPoint(0)
    end1 = curve1.GetEndPoint(1)
    start2 = curve2.GetEndPoint(0)
    end2 = curve2.GetEndPoint(1)

    d1x, d1y, d1z = end1.X - start1.X, end1.Y - start1.Y, end1.Z - start1.Z
    d2x, d2y, d2z = end2.X - start2.X, end2.Y - start2.Y, end2.Z - start2.Z

    w0x, w0y, w0z = start1.X - start2.X, start1.Y - start2.Y, start1.Z - start2.Z

    a = d1x * d1x + d1y * d1y + d1z * d1z
    b = d1x * d2x + d1y * d2y + d1z * d2z
    c = d2x * d2x + d2y * d2y + d2z * d2z
    d = d1x * w0x + d1y * w0y + d1z * w0z
    e = d2x * w0x + d2y * w0y + d2z * w0z

    denom = a * c - b * b
    if denom == 0:
        return float('inf'), None
    
    t = (b * e - c * d) / denom
    u = (a * e - b * d) / denom

    closest_point_line1 = (start1.X + t * d1x, start1.Y + t * d1y, start1.Z + t * d1z)
    closest_point_line2 = (start2.X + u * d2x, start2.Y + u * d2y, start2.Z + u * d2z)

    distance = ((closest_point_line1[0] - closest_point_line2[0]) ** 2 +
                (closest_point_line1[1] - closest_point_line2[1]) ** 2 +
                (closest_point_line1[2] - closest_point_line2[2]) ** 2) ** 0.5

    return distance, XYZ(closest_point_line1[0], closest_point_line1[1], closest_point_line1[2])

# Funzione per ottenere i connettori di un tubo
def get_connectors(pipe):
    try:
        return pipe.ConnectorManager.Connectors
    except:
        return None

# Funzione per trovare i due connettori non connessi piu vicini
def find_closest_connectors(pipe1, pipe2):
    connectors1 = get_connectors(pipe1)
    connectors2 = get_connectors(pipe2)
    
    if connectors1 is None or connectors2 is None:
        return None, None

    min_distance = float('inf')
    closest_pair = None

    for connector1 in connectors1:
        if not connector1.IsConnected:
            for connector2 in connectors2:
                if not connector2.IsConnected:
                    distance = connector1.Origin.DistanceTo(connector2.Origin)
                    if distance < min_distance:
                        min_distance = distance
                        closest_pair = (connector1, connector2)

    return closest_pair

# Funzione per verificare se una linea e allineata agli assi principali
def is_aligned_to_axes(line):
    direction = line.Direction
    return (abs(direction.X) > 0.99 and abs(direction.Y) < 0.01 and abs(direction.Z) < 0.01) or \
           (abs(direction.Y) > 0.99 and abs(direction.X) < 0.01 and abs(direction.Z) < 0.01) or \
           (abs(direction.Z) > 0.99 and abs(direction.X) < 0.01 and abs(direction.Y) < 0.01)

# Funzione per spostare rigidamente un tubo solo sull'asse X o Y
def translate_pipe(pipe, translation_vector):
    location = pipe.Location
    if isinstance(location, LocationCurve):
        # Controlliamo quale tra X e Y ha il valore di traslazione maggiore
        if abs(translation_vector.X) > abs(translation_vector.Y):
            # Traslazione lungo l'asse X
            translation_vector = XYZ(translation_vector.X, 0, 0)
        else:
            # Traslazione lungo l'asse Y
            translation_vector = XYZ(0, translation_vector.Y, 0)
        location.Move(translation_vector)

# Informa l'utente prima della selezione
TaskDialog.Show('Selezione', 'Seleziona il tubo di riferimento.')

# Seleziona il tubo di riferimento in Revit
try:
    sel = uidoc.Selection
    ref_pipe_ref = sel.PickObject(ObjectType.Element, "Seleziona il tubo di riferimento.")
    Pipe_2 = doc.GetElement(ref_pipe_ref.ElementId)

    # Mostra il messaggio per la seconda selezione
    TaskDialog.Show('Selezione', 'Seleziona il tubo da modificare.')

    # Seleziona il tubo da modificare
    mod_pipe_ref = sel.PickObject(ObjectType.Element, "Seleziona il tubo da modificare.")
    Pipe_1 = doc.GetElement(mod_pipe_ref.ElementId)

except Exception as e:
    TaskDialog.Show('Errore', 'Errore nella selezione degli elementi: {}'.format(e))
    script.exit()

# Ottieni le curve (linee) dei tubi
Linea_Riferimento = Pipe_2.Location.Curve
Linea_Modificare = Pipe_1.Location.Curve

# Verifica se i tubi sono allineati lungo uno degli assi principali
if is_aligned_to_axes(Linea_Riferimento) and is_aligned_to_axes(Linea_Modificare):
    # Se sono allineati, calcola il vettore di traslazione
    translation_vector = Linea_Riferimento.GetEndPoint(0) - Linea_Modificare.GetEndPoint(0)
    
    # Inizia la transazione per la traslazione rigida
    t = Transaction(doc, 'Traslazione rigida dei tubi')
    t.Start()

    # Sposta il tubo lungo l'asse X o Y (Z resta invariato)
    translate_pipe(Pipe_1, translation_vector)

    # Termina la transazione
    t.Commit()

    # Inizia una seconda transazione per creare il raccordo (elbow)
    t2 = Transaction(doc, 'Creazione del raccordo')
    t2.Start()

    # Trova i connettori piu vicini tra i due tubi
    connector1, connector2 = find_closest_connectors(Pipe_1, Pipe_2)

    # Se ci sono connettori non connessi, crea un raccordo (fitting)
    if connector1 and connector2:
        doc.Create.NewElbowFitting(connector1, connector2)
        TaskDialog.Show('Successo', "Fitting creato tra {0} e {1}".format(Pipe_1.Id, Pipe_2.Id))

    # Termina la seconda transazione
    t2.Commit()

else:
    # Se non sono allineati, calcola la distanza e il punto piu vicino tra le due curve
    Distanza, PuntoVicino = ClosestPointAndDDistance(Linea_Riferimento, Linea_Modificare)

    # Crea una nuova linea utilizzando il punto piu lontano e il punto piu vicino
    LineaNuova = Line.CreateBound(FarPoint(PuntoVicino, Linea_Modificare), PuntoVicino)

    # Inizia la prima transazione per l'allineamento
    t = Transaction(doc, 'Allineamento e connessione dei tubi')
    t.Start()

    # Aggiorna la curva del primo tubo
    Pipe_1.Location.Curve = LineaNuova

    # Termina la prima transazione
    t.Commit()

    # Inizia una seconda transazione per creare il raccordo (elbow)
    t2 = Transaction(doc, 'Creazione del raccordo')
    t2.Start()

    # Trova i connettori piu vicini tra i due tubi
    connector1, connector2 = find_closest_connectors(Pipe_1, Pipe_2)

    # Se ci sono connettori non connessi, crea un raccordo (fitting)
    if connector1 and connector2:
        doc.Create.NewElbowFitting(connector1, connector2)
        TaskDialog.Show('Successo', "Fitting creato tra {0} e {1}".format(Pipe_1.Id, Pipe_2.Id))

    # Termina la seconda transazione
    t2.Commit()