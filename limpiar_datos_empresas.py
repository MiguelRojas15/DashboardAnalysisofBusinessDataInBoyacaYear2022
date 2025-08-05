import pandas as pd
import numpy as np

# Archivo de Excel con los datos originales
archivo = "PROYECTO BOYACÁ EMPRESAS.xlsx"

# Cargar la hoja con los datos
df = pd.read_excel(archivo, sheet_name='Datos_Empresas_Boyaca_2022_ENRI')

# Limpiar datos
df['ProductoElaborado'] = df['ProductoElaborado'].replace('Sin datos', np.nan)
df['ProgramaVinculado'] = df['ProgramaVinculado'].str.replace(';x|', ';', regex=True)
df['AcompanamientoRecibido'] = df['AcompanamientoRecibido'].fillna('No reportado')
df['Ventas mensuales (Millones)'] = pd.to_numeric(df['Ventas mensuales (Millones)'], errors='coerce')

# Guardar archivo limpio
df.to_excel("PROYECTO_BOYACA_EMPRESAS_LIMPIO.xlsx", index=False)

print(" Limpieza completada. Archivo guardado como PROYECTO_BOYACA_EMPRESAS_LIMPIO.xlsx")