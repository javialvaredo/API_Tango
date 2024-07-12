import requests
import pandas as pd
from typing import List, Dict
from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess

class ApiClient:
    def __init__(self, base_url: str, api_key: str, company_id: str):
        self.base_url = base_url
        self.headers = {
            'ApiAuthorization': api_key,
            'Company': company_id
        }

    def get_data(self, process: str, from_date: str, to_date: str, page_size: str) -> List[Dict]:
        params = {
            'process': process,
            'fromDate': from_date,
            'toDate': to_date,
            'pageSize': page_size,
            'pageIndex': '0',
            'customQuery': '0'
        }
        try:
            response = requests.get(self.base_url, headers=self.headers, params=params)
            response.raise_for_status()
            data = response.json()
            return data.get('resultData', {}).get('list', [])
        except requests.exceptions.RequestException as e:
            print(f'Error en la solicitud: {e}')
            return []
        except ValueError as e:
            print('Error al analizar JSON:', e)
            return []

class ExcelConverter:
    def __init__(self, data: List[Dict]):
        self.data = data
        self.last_file_path = None  # Para almacenar la ruta del último archivo generado

    def to_excel(self, file_name: str):
        if not self.data:
            print('No hay datos para convertir a Excel.')
            return

        try:
            df = pd.DataFrame(self.data)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name_with_timestamp = f"{file_name}_{timestamp}.xlsx"
            output_directory = os.path.join(os.getcwd(), 'pedidos')
            if not os.path.exists(output_directory):
                os.makedirs(output_directory)
            file_path = os.path.join(output_directory, file_name_with_timestamp)
            df.to_excel(file_path, index=False)
            self.last_file_path = file_path  # Guardar la ruta del archivo generado
            print(f'Archivo Excel guardado en: {file_path}')
        except Exception as e:
            print(f'Error al guardar el archivo Excel: {e}')


# Clase para la interfaz gráfica
class VentanaGrafica(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Entrada de Parámetros  -  API TANGO")
        self.geometry("400x300")

        # Valores por defecto
        self.default_values = {
            'proceso': '10081',
            'fecha desde': '01/01/2024',
            'fecha hasta': '30/06/2024',
            'tamaño de página': '100'
        }

        # Etiquetas y campos de entrada
        self.crear_campos()

        # Botón de envío
        self.boton_enviar = tk.Button(self, text="Enviar", command=self.enviar)
        self.boton_enviar.pack(pady=10)

    def crear_campos(self):
        self.campos = {}
        for campo, valor_defecto in self.default_values.items():
            frame = tk.Frame(self)
            frame.pack(pady=5)
            label = tk.Label(frame, text=campo)
            label.pack(side=tk.LEFT)
            entry = tk.Entry(frame)
            entry.insert(0, valor_defecto)  # Insertar valor por defecto
            entry.pack(side=tk.RIGHT)
            self.campos[campo] = entry

    def enviar(self):
        # Obtener los valores de los campos
        process = self.campos['proceso'].get()
        from_date = self.campos['fecha desde'].get()
        to_date = self.campos['fecha hasta'].get()
        page_size = self.campos['tamaño de página'].get()

        # Validación básica de entradas
        if not all([process, from_date, to_date, page_size]):
            messagebox.showerror("Error", "Todos los campos deben estar llenos")
            return
        
          # Verificar que process y page_size sean numéricos
        if not (process.isdigit() and page_size.isdigit()):
            messagebox.showerror("Error", "Los campos 'Process' y 'Page Size' deben ser numéricos")
            return False
        
        # Verificar que las fechas estén en el formato dd/mm/aaaa
        date_format = "%d/%m/%Y"
        try:
            datetime.strptime(from_date, date_format)
            datetime.strptime(to_date, date_format)
        except ValueError:
            messagebox.showerror("Error", "Las fechas deben estar en el formato dd/mm/aaaa")
            return False

            

        # Instanciar el cliente API
        base_url = 'http://cv-tango:17000/Api/GetApiLiveQueryData'
        api_key = '2f34f26e-c228-4dd2-af3b-616da1112a1e'
        company_id = '4'
        client = ApiClient(base_url, api_key, company_id)

        # Obtener los datos
        datos = client.get_data(process, from_date, to_date, page_size)
        print('datos: ', type(datos))

        # Instanciar el convertidor de Excel
        converter = ExcelConverter(datos)

        # Guardar los datos en un archivo Excel en la subcarpeta 'pedidos'
        output_file = 'Pedidos'
        converter.to_excel(output_file)

        if converter.last_file_path:
            messagebox.showinfo("Completado", f"Los datos se han guardado en un archivo Excel y se han abierto: {converter.last_file_path}")
            self.abrir_archivo(converter.last_file_path)

    def abrir_archivo(self, file_path):
        try:
            if os.name == 'nt':  # Si es Windows
                os.startfile(file_path)
            elif os.name == 'posix':  # Si es Unix/Linux/Mac
                subprocess.run(['open', file_path])
            else:
                print('No se puede abrir el archivo automáticamente en este sistema operativo.')
        except Exception as e:
            print(f'Error al abrir el archivo Excel: {e}')
            messagebox.showerror("Error", f"Error al abrir el archivo: {e}")

if __name__ == '__main__':
    app = VentanaGrafica()
    app.mainloop()
