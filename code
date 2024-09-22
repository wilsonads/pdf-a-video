import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import re

def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return ruta_archivo

def procesar_archivo(ruta_archivo):
    # Diccionario de marcas con patrones regex para variaciones comunes
    marcas_variaciones = {
        "ENSURE": r"ENSURE|ENSUR|ENSR",
        "GLUCERNA": r"GLUCERNA|GLUCERN|GLUC",
        "PEDIASURE": r"PEDIASURE|PEDSURE|PEDIASRE|PEDIAUSRE|PEDIAUSURE|PEDIASUR",
        "SIMILAC": r"SIMILAC|SMILAC|SIMILA",
        "PEDIALYTE": r"PEDIALYTE|PEDIALYT|PEDIALY",
        "OTRO": r"OTRO"
    }

    try:
        workbook = openpyxl.load_workbook(ruta_archivo)
    except PermissionError:
        messagebox.showerror("Error", "No se pudo abrir el archivo. Asegúrate de que no esté abierto en otro programa y que tengas permisos para acceder a él.")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al abrir el archivo: {str(e)}")
        return
    
    sheet = workbook.active
    
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):  # Leer desde la columna 1 (columna A)
        texto_celda = str(row[0].value).upper() if row[0].value else ""
        marcas_encontradas = []
        
        for marca, patron in marcas_variaciones.items():
            # Buscar coincidencias con el patrón de cada marca
            if re.search(patron, texto_celda):
                marcas_encontradas.append(marca)
        
        if marcas_encontradas:
            # Colocar resultados en la columna 2 (columna B)
            row[0].offset(column=1).value = ", ".join(marcas_encontradas)
    
    try:
        workbook.save(ruta_archivo)
        messagebox.showinfo("Éxito", "Procesamiento completado.")
    except PermissionError:
        messagebox.showerror("Error", "No se pudo guardar el archivo. Asegúrate de que no esté abierto en otro programa y que tengas permisos para escribir en él.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar el archivo: {str(e)}")

# Ejecución principal
ruta_archivo = seleccionar_archivo()
if ruta_archivo:
    procesar_archivo(ruta_archivo)
else:
    messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
