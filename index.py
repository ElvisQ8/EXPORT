import streamlit as st
import pandas as pd
import openpyxl
import requests
from io import BytesIO

# Función para cargar el archivo certificado
def cargar_certificado(file):
    # Cargar los datos desde la fila 27 hasta la 127, considerando la fila 26 como cabecera
    df = pd.read_excel(file, sheet_name=0, header=25, skiprows=0, nrows=101)
    st.write(df.head())  # Mostrar las primeras filas para verificar las columnas
    return df

# Función para cargar la plantilla desde GitHub
def cargar_plantilla_github():
    url = "https://github.com/tu_usuario/tu_repositorio/raw/main/plantilla_export.xlsx"
    response = requests.get(url)
    if response.status_code == 200:
        plantilla = openpyxl.load_workbook(BytesIO(response.content))
        return plantilla
    else:
        st.error("No se pudo cargar el archivo plantilla_export.xlsx desde GitHub.")
        return None

# Función para aplicar filtros y copiar los datos filtrados a la plantilla
def procesar_datos(certificado, plantilla):
    # Filtrar y preparar los datos para las hojas correspondientes
    hoja_o = plantilla["O"]
    hoja_dp = plantilla["DP"]
    hoja_std = plantilla["STD"]
    
    # Filtrar para la hoja "O" (sin "DEND" ni "DSTD")
    datos_o = certificado[~certificado['O'].str.contains('DEND|DSTD', na=False)]
    for i, row in datos_o.iterrows():
        hoja_o.append([row['A'], row['B'], row['C'], row['D'], row['E'], row['F'], row['G'], row['H'],
                       row['I'], row['J'], row['K'], row['L'], row['N'], None, row['Q'], None, row['R']])

    # Filtrar para la hoja "DP" (con "DEND")
    datos_dp = certificado[certificado['O'].str.contains('DEND', na=False)]
    for i, row in datos_dp.iterrows():
        hoja_dp.append([row['A'], row['B'], row['C'], row['D'], row['E'], row['F'], row['G'], row['H'],
                       row['I'], row['J'], row['K'], row['O'], row['N'], row['Q'], None, row['R']])

    # Filtrar para la hoja "STD" (con "DSTD")
    datos_std = certificado[certificado['O'].str.contains('DSTD', na=False)]
    for i, row in datos_std.iterrows():
        hoja_std.append([row['A'], row['E'], row['G'], row['H'], row['I'], row['J'], row['K'], row['PECLSTDEN02'],
                        row['N'], row['Q'], None, row['R']])

    # Guardar la plantilla con los nuevos datos
    plantilla.save("plantilla_export_modificada.xlsx")

# Función para exportar una hoja a CSV
def exportar_csv(hoja_nombre, plantilla):
    hoja = plantilla[hoja_nombre]
    df = pd.DataFrame(hoja.values)
    df.to_csv(f"{hoja_nombre}.csv", index=False)

# Interfaz de usuario en Streamlit
st.title("Procesar Certificado y Exportar Datos")

# Subir archivo certificado
uploaded_file = st.file_uploader("Sube el archivo certificado.xlsx", type=["xlsx"])

if uploaded_file is not None:
    # Cargar y procesar el archivo certificado
    certificado = cargar_certificado(uploaded_file)
    plantilla = cargar_plantilla_github()

    if plantilla:
        # Procesar datos y copiarlos a la plantilla
        procesar_datos(certificado, plantilla)
    
        st.success("Los datos han sido procesados y copiados exitosamente.")
    
        # Generar botones para exportar a CSV
        if st.button('Exportar hoja O a CSV'):
            exportar_csv('O', plantilla)
            st.success("Hoja 'O' exportada a CSV.")
    
        if st.button('Exportar hoja DP a CSV'):
            exportar_csv('DP', plantilla)
            st.success("Hoja 'DP' exportada a CSV.")
    
        if st.button('Exportar hoja STD a CSV'):
            exportar_csv('STD', plantilla)
            st.success("Hoja 'STD' exportada a CSV.")
