import tkinter as tk
from tkinter import ttk, messagebox
from file_manager import FileManager
from data_handler import DataHandler
from PIL import Image, ImageTk

class GUI:
    def __init__(self):
        self.excel_file = None
        self.hojas = None
        self.hoja_seleccionada = None

        self.ventana = tk.Tk()
        self.ventana.title("Procesador de Facturas")
        self.ventana.geometry("600x600")
        self.ventana.configure(bg='#f8f8f8')

        self.crear_componentes()
        self.ventana.mainloop()

    def crear_componentes(self):
        # Logo
        try:
            img = Image.open(r"C:\Users\CONGRESO\Desktop\programa\imagenes\logo2.jpeg")
            logo = ImageTk.PhotoImage(img)
            logo_label = tk.Label(self.ventana, image=logo, bg='#f8f8f8')
            logo_label.image = logo  # Mantener referencia
            logo_label.pack(pady=10)
        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo cargar el logo: {e}")

        # Barra de progreso
        self.progress_label = tk.Label(self.ventana, text="Listo para comenzar", bg='#f8f8f8')
        self.progress_label.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self.ventana, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Etiqueta para mostrar la cantidad total de links en la hoja seleccionada
        self.cantidad_links_label = tk.Label(self.ventana, text="Cantidad total de links: 0", bg='#f8f8f8')
        self.cantidad_links_label.pack(pady=5)

        # Cuadro de texto para ingresar la cantidad de links
        etiqueta_links = tk.Label(self.ventana, text="Cantidad de links a procesar:", bg='#f8f8f8')
        etiqueta_links.pack()
        self.entrada_links = tk.Entry(self.ventana, width=10, justify='center')
        self.entrada_links.pack(pady=5)

        # Botón para cargar archivo Excel
        boton_cargar_excel = tk.Button(self.ventana, text="Cargar Excel", command=self.cargar_archivo, width=20, height=2, bg='navy', fg='white')
        boton_cargar_excel.pack(pady=10)

        # Botón para seleccionar hoja del Excel
        boton_seleccionar_hoja = tk.Button(self.ventana, text="Seleccionar Hoja", command=self.seleccionar_hoja, width=20, height=2, bg='blue', fg='white')
        boton_seleccionar_hoja.pack(pady=10)

        # Etiqueta para mostrar el total de links escaneados
        self.total_label = tk.Label(self.ventana, text="Total de links escaneados: 0", bg='#f8f8f8')
        self.total_label.pack(pady=5)

        # Botón para iniciar el proceso
        boton_iniciar_proceso = tk.Button(self.ventana, text="Iniciar Proceso", command=self.iniciar_proceso, width=20, height=2, bg='red', fg='white')
        boton_iniciar_proceso.pack(pady=10)

    def cargar_archivo(self):
        self.excel_file, self.hojas = FileManager.cargar_excel()
        if self.hojas:
            messagebox.showinfo("Hojas", f"El archivo tiene las siguientes hojas: {', '.join(self.hojas)}")

    def seleccionar_hoja(self):
        if self.hojas is not None:
            seleccion_hoja = tk.Toplevel(self.ventana)
            seleccion_hoja.title("Selecciona una hoja")
            seleccion_hoja.geometry("300x200")

            lista_hojas = ttk.Combobox(seleccion_hoja, values=self.hojas, state="readonly")
            lista_hojas.pack(pady=20)

            def seleccionar():
                hoja = lista_hojas.get()
                if hoja:
                    self.hoja_seleccionada = self.excel_file.parse(hoja)
                    self.cantidad_links_label.config(text=f"Cantidad total de links: {len(self.hoja_seleccionada)}")
                    seleccion_hoja.destroy()

            boton_seleccionar = tk.Button(seleccion_hoja, text="Seleccionar", command=seleccionar)
            boton_seleccionar.pack(pady=10)

    def iniciar_proceso(self):
        cantidad_links = int(self.entrada_links.get())
        if self.hoja_seleccionada is not None and cantidad_links > 0:
            datos_procesados = DataHandler.procesar_datos(self.hoja_seleccionada, cantidad_links, self.progress_bar, self.progress_label, self.total_label)
            if not datos_procesados.empty:
                FileManager.guardar_excel(datos_procesados)
            self.progress_label.config(text="Proceso completado")
        else:
            self.progress_label.config(text="Error en el archivo Excel o en la cantidad de links")

# Iniciar la aplicación
if __name__ == "__main__":
    GUI()
