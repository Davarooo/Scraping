import requests  # Obtener páginas web (HTTP requests)
from bs4 import BeautifulSoup  # Analizar y extraer datos de HTML/XML
import pandas as pd  # Manipular datos en tablas y exportar a Excel/CSV
from datetime import datetime  # Trabajar con fechas y horas
import os  # Interactuar con el sistema operativo (archivos, rutas)
import pyautogui  # Automatizar interacciones con mouse/teclado
import time  # Controlar tiempos de espera y retardos

# 1. Obtener el precio del dólar desde la web
url = 'https://www.dolar-colombia.com/'
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

# Buscar el valor dentro del <span> con clase que contiene 'exchange-rate'
valor_span = soup.find('span', class_=lambda x: x and 'exchange-rate' in x)

# Validar si se encontró el dato
if valor_span is None:
    print("❌ No se encontró el valor del dólar. Verifica la estructura de la página.")
    exit()

# Limpiar el texto del valor
valor = valor_span.text.strip().replace('$', '').replace('.', '').replace(',', '.')

# Preparar los datos con la fecha actual
hoy = datetime.now().strftime('%Y-%m-%d')
datos = {'Fecha': hoy, 'Valor USD': float(valor)}

# 2. Guardar los datos en un archivo Excel
archivo = 'dolar_historico.xlsx'

if os.path.exists(archivo):
    df = pd.read_excel(archivo)
    nuevo_df = pd.DataFrame([datos])
    df = pd.concat([df, nuevo_df], ignore_index=True)
else:
    df = pd.DataFrame([datos])

df.to_excel(archivo, index=False)

print(f"✅ Precio del dólar guardado: {hoy} - ${valor}")

# 3. Automatización para abrir el archivo con pyautogui

time.sleep(2)  # Espera breve antes de mover el mouse

# Abrir el menú inicio (Windows)
pyautogui.press('win')
time.sleep(1)

# Escribir el nombre del archivo (debe estar en el Escritorio o en una ruta accesible)
pyautogui.write('dolar_historico.xlsx')
time.sleep(1)

# Abrir el archivo
pyautogui.press('enter')

print("🖱️ Excel abierto automáticamente con pyautogui.")
