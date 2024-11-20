import pandas as pd
import os
from tkinter import filedialog, messagebox

class FileManager:
    @staticmethod
    def cargar_excel():
        archivo = filedialog.askopenfilename(title="Selecciona un archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
        if archivo:
            try:
                excel_file = pd.ExcelFile(archivo)
                hojas = excel_file.sheet_names
                return excel_file, hojas
            except Exception as e:
                messagebox.showerror("Error", f"Error al leer el archivo Excel: {e}")
                return None, None
        else:
            messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
            return None, None

    @staticmethod
    def guardar_excel(dataframe):
        ruta_guardado = r"C:\Users\CONGRESO\Desktop\programa\excel nuevo"
        if not os.path.exists(ruta_guardado):
            os.makedirs(ruta_guardado)

        archivo_guardar = os.path.join(ruta_guardado, "resultado facturas.xlsx")
        dataframe.to_excel(archivo_guardar, index=False)
        messagebox.showinfo("Éxito", f"Archivo guardado en: {archivo_guardar}")
