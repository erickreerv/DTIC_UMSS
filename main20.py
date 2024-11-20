import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import requests
from bs4 import BeautifulSoup
from collections import Counter

# Función para obtener datos de la URL
def obtener_datos(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            tables = soup.find_all('table')
            for table in tables:
                rows = table.find_all('tr')
                if rows:
                    headers = [header.text.strip() for header in rows[0].find_all('th')]
                    if 'NRO. DE FACTURA' in headers and 'MONTO' in headers:
                        index_factura = headers.index('NRO. DE FACTURA')
                        index_monto = headers.index('MONTO')
                        datos = []
                        for row in rows[1:]:
                            data = [cell.text.strip() for cell in row.find_all('td')]
                            if len(data) > max(index_factura, index_monto):
                                numero_factura = data[index_factura]
                                monto = data[index_monto].replace('$', '').replace(',', '').strip()
                                estado = 'Válido' if numero_factura.isdigit() and int(numero_factura) > 0 else 'No Válido'
                                datos.append((url, int(numero_factura), float(monto), estado))
                        return datos
        return None
    except Exception:
        return None

# Función para cargar el archivo Excel y obtener las hojas
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

# Función para procesar los datos del Excel
def procesar_datos(dataframe, cantidad_links, progress_bar, progress_label, total_label):
    datos_procesados = []
    total_links = min(cantidad_links, len(dataframe))
    total_escaneados = 0

    if 'Link' not in dataframe.columns:
        messagebox.showwarning("Advertencia", "No se encontró la columna 'Link' en el archivo Excel.")
        return pd.DataFrame()

    for index, row in dataframe.iterrows():
        if index >= total_links:
            break
        url = row['Link']
        if not url or pd.isna(url):
            datos_procesados.append((index + 1, "Fila sin datos", "Fila sin datos", "Fila sin datos", "Fila sin datos"))
            continue

        progress_label.config(text=f"Procesando link {index + 1} de {total_links}...")
        progress_bar['value'] = (index + 1) / total_links * 100
        progress_bar.update()

        datos = obtener_datos(url)
        if datos:
            for dato in datos:
                datos_procesados.append((index + 1,) + dato)
            total_escaneados += len(datos)
        else:
            datos_procesados.append((index + 1, url, 'Error', 'Error', 'Error'))

        total_label.config(text=f"Total de links escaneados: {total_escaneados}")
        total_label.update()

    df_resultado = pd.DataFrame(datos_procesados, columns=['Index', 'Link', 'Nro Factura', 'Monto', 'Estado'])

    # Verificar duplicados y marcar en el DataFrame
    duplicados = df_resultado['Nro Factura'].duplicated(keep=False)
    df_resultado['Duplicado'] = duplicados.replace({True: 'Duplicado', False: 'Único'})

    return df_resultado

# Función para guardar el nuevo archivo Excel
def guardar_excel(dataframe):
    ruta_guardado = r"C:\Users\Usuario\Downloads\facturas agosto"
    if not os.path.exists(ruta_guardado):
        os.makedirs(ruta_guardado)

    archivo_guardar = os.path.join(ruta_guardado, "resultado_facturas.xlsx")
    
    # Resaltar filas duplicadas con colores claros diferentes
    with pd.ExcelWriter(archivo_guardar, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Resultados')
        
        workbook = writer.book
        worksheet = writer.sheets['Resultados']

        # Identificar duplicados por 'Nro Factura' y agruparlos
        duplicados = dataframe[dataframe['Duplicado'] == 'Duplicado']
        grupos = duplicados.groupby('Nro Factura').groups

        # Definir colores claros
        colores = ['#FFCCCB', '#FFFFCC', '#CCFFCC', '#CCE5FF', '#FFCCE5']

        # Aplicar formato de color por grupo
        for i, (factura, indices) in enumerate(grupos.items()):
            color = colores[i % len(colores)]  # Ciclar colores si hay más grupos
            formato = workbook.add_format({'bg_color': color})

            for row_num in indices:
                worksheet.set_row(row_num + 1, None, formato)  # +1 para evitar el encabezado

    messagebox.showinfo("Éxito", f"Archivo guardado en: {archivo_guardar}")

# Función para verificar duplicados
def verificar_duplicados(dataframe):
    if 'Nro Factura' in dataframe.columns:
        duplicados = dataframe['Nro Factura'].value_counts()
        duplicados = duplicados[duplicados > 1]
        if not duplicados.empty:
            duplicados_texto = '\n'.join([f'Factura {num}: {count} veces' for num, count in duplicados.items()])
            messagebox.showinfo("Duplicados Encontrados", f"Los siguientes números de factura están duplicados:\n{duplicados_texto}")
        else:
            messagebox.showinfo("Sin Duplicados", "No se encontraron duplicados.")
    else:
        messagebox.showwarning("Advertencia", "No se encontró la columna 'Nro Factura'.")

# Función para mostrar los totales y cantidad de links
def actualizar_datos_links(dataframe, cantidad_links_label):
    total_links = len(dataframe)
    cantidad_links_label.config(text=f"Cantidad total de links: {total_links}")
    return total_links

# Función que gestiona el proceso completo
def ejecutar_proceso(progress_bar, progress_label, hoja_seleccionada, cantidad_links, total_label):
    if hoja_seleccionada is not None and cantidad_links > 0:
        datos_procesados = procesar_datos(hoja_seleccionada, cantidad_links, progress_bar, progress_label, total_label)
        if not datos_procesados.empty:
            guardar_excel(datos_procesados)
            verificar_duplicados(datos_procesados)
        progress_label.config(text="Proceso completado")
    else:
        progress_label.config(text="Error en el archivo Excel o en la cantidad de links")

# Función para seleccionar la hoja
def seleccionar_hoja(excel_file, hojas, cantidad_links_label):
    seleccion_hoja = tk.Toplevel()
    seleccion_hoja.title("Selecciona una hoja")
    seleccion_hoja.geometry("300x200")

    lista_hojas = ttk.Combobox(seleccion_hoja, values=hojas, state="readonly")
    lista_hojas.pack(pady=20)

    def seleccionar():
        hoja = lista_hojas.get()
        if hoja:
            global hoja_seleccionada
            hoja_seleccionada = excel_file.parse(hoja)
            cantidad_links_label.config(text=f"Cantidad total de links: {len(hoja_seleccionada)}")
            seleccion_hoja.destroy()

    boton_seleccionar = tk.Button(seleccion_hoja, text="Seleccionar", command=seleccionar)
    boton_seleccionar.pack(pady=10)

# Crear la interfaz gráfica
def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Procesador de Facturas")
    ventana.geometry("600x600")
    ventana.configure(bg='#f8f8f8')

    global excel_file, hojas, hoja_seleccionada
    excel_file = None
    hojas = None
    hoja_seleccionada = None

    # Logo
    try:
        logo = tk.PhotoImage(file=r"C:\Users\Usuario\Desktop\programa\logoumss.png")
        logo_label = tk.Label(ventana, image=logo, bg='#f8f8f8')
        logo_label.image = logo
        logo_label.pack(pady=10)
    except Exception as e:
        messagebox.showwarning("Advertencia", f"No se pudo cargar el logo: {e}")

    # Barra de progreso
    progress_label = tk.Label(ventana, text="Listo para comenzar", bg='#f8f8f8')
    progress_label.pack(pady=10)

    ventana.progress_bar = ttk.Progressbar(ventana, orient="horizontal", length=300, mode="determinate")
    ventana.progress_bar.pack(pady=10)

    cantidad_links_label = tk.Label(ventana, text="Cantidad total de links: 0", bg='#f8f8f8')
    cantidad_links_label.pack(pady=5)

    etiqueta_links = tk.Label(ventana, text="Cantidad de links a procesar:", bg='#f8f8f8')
    etiqueta_links.pack()
    entrada_links = tk.Entry(ventana, width=10, justify='center')
    entrada_links.pack(pady=5)

    def cargar_archivo():
        global excel_file, hojas
        excel_file, hojas = cargar_excel()
        if hojas:
            messagebox.showinfo("Hojas", f"El archivo tiene las siguientes hojas: {', '.join(hojas)}")

    boton_cargar_excel = tk.Button(ventana, text="Cargar Excel", command=cargar_archivo, width=20, height=2, bg='navy', fg='white')
    boton_cargar_excel.pack(pady=10)

    boton_seleccionar_hoja = tk.Button(ventana, text="Seleccionar Hoja", command=lambda: seleccionar_hoja(excel_file, hojas, cantidad_links_label), width=20, height=2, bg='blue', fg='white')
    boton_seleccionar_hoja.pack(pady=10)

    total_label = tk.Label(ventana, text="Total de links escaneados: 0", bg='#f8f8f8')
    total_label.pack(pady=5)

    boton_iniciar_proceso = tk.Button(ventana, text="Iniciar Proceso", command=lambda: ejecutar_proceso(ventana.progress_bar, progress_label, hoja_seleccionada, int(entrada_links.get()), total_label), width=20, height=2, bg='red', fg='white')
    boton_iniciar_proceso.pack(pady=10)

    ventana.mainloop()

# Ejecutar la interfaz gráfica
crear_interfaz()
