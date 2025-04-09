import ifcopenshell
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import os  # Aggiunto per gestire il percorso della cartella corrente


current_folder = os.path.dirname(__file__)  # Ottieni la cartella dello script


def apri_file_ifc():
    root = tk.Tk()
    root.withdraw()  # Nascondi la finestra principale di Tkinter
    file_path = filedialog.askopenfilename(
        title="Seleziona il file IFC",
        filetypes=(("IFC files", "*.ifc"), ("All files", "*.*"))
    )
    return file_path

def estrai_tutti_elementi(file_ifc):
    # Apri il file IFC usando ifcopenshell
    model = ifcopenshell.open(file_ifc)
    elementi = []
    # Estrae tutti gli oggetti derivanti da IfcElement
    for obj in model.by_type("IfcElement"):
        name = getattr(obj, 'Name', None).split(":")[0] + " - " + getattr(obj, 'Name', None).split(":")[1]
        guid = getattr(obj, 'GlobalId', None)
        elementid = obj.Tag
        if guid:
            elementi.append((name, obj.is_a(), guid, elementid))
    return elementi

def main():
    file_ifc = apri_file_ifc()
    if not file_ifc:
        messagebox.showerror("Errore", "Nessun file IFC selezionato.")
        return

    elementi = estrai_tutti_elementi(file_ifc)
    
    if not elementi:
        messagebox.showinfo("Informazione", "Nessun elemento trovato nel file IFC.")
    else:
        # Prepara un report testuale con la corrispondenza "Classe IFC - GUID"
        #risultati = "\n".join([f"{tipo}: {guid}" for tipo, guid in elementi])
        # Visualizza il report in una finestra di messaggio
        #messagebox.showinfo("Elementi presenti nel file IFC", risultati)
        
        # Salva anche i risultati in un file CSV nella cartella corrente
        csv_path = os.path.join(current_folder, "risultati_ifc.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Nome Elemento","Tipo IFC", "GUID","ElementId"])
            for nome, tipo, guid, elementid in elementi:
                writer.writerow([nome, tipo, guid, elementid])
        print(f"Risultati salvati in '{csv_path}'")

if __name__ == "__main__":
    main()
