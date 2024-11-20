import pandas as pd
import requests
from bs4 import BeautifulSoup
from tkinter import messagebox  # Importar messagebox

class DataHandler:
    @staticmethod
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
                            for row in rows[1:]:
                                data = [cell.text.strip() for cell in row.find_all('td')]
                                if len(data) > max(index_factura, index_monto):
                                    numero_factura = data[index_factura]
                                    monto = data[index_monto].replace('$', '').replace(',', '').strip()
                                    estado = 'Válido' if numero_factura.isdigit() and int(numero_factura) > 0 else 'No Válido'
                                    return (url, int(numero_factura), float(monto), estado)
            return None
        except Exception:
            return None

    @staticmethod
    def procesar_datos(dataframe, cantidad_links, progress_bar, progress_label, total_label):
        datos_procesados = []
        total_links = min(cantidad_links, len(dataframe))
        total_escaneados = 0  # Contador para los links escaneados

        if 'Link' not in dataframe.columns:
            messagebox.showwarning("Advertencia", "No se encontró la columna 'Link' en el archivo Excel.")
            return pd.DataFrame()

        for index, row in dataframe.iterrows():
            if index >= total_links:
                break
            url = row['Link']
            if not url or pd.isna(url):  # Verificar si la fila está vacía
                datos_procesados.append((index + 1, "Fila sin datos", "Fila sin datos", "Fila sin datos", "Fila sin datos"))
                continue

            progress_label.config(text=f"Procesando link {index + 1} de {total_links}...")
            progress_bar['value'] = (index + 1) / total_links * 100  # Actualiza la barra de progreso
            progress_bar.update()

            datos = DataHandler.obtener_datos(url)
            if datos:
                datos_procesados.append((index + 1,) + datos)
                total_escaneados += 1
            else:
                datos_procesados.append((index + 1, url, 'Error', 'Error', 'Error'))

            # Actualizar el total de links escaneados
            total_label.config(text=f"Total de links escaneados: {total_escaneados}")
            total_label.update()

        df_resultado = pd.DataFrame(datos_procesados, columns=['Index', 'Link', 'Nro Factura', 'Monto', 'Estado'])
        return df_resultado

