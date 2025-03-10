import streamlit as st
import pandas as pd
import openpyxl
import io

# Ruta local del archivo plantilla
PLANTILLA_PATH = "plantilla_export.xlsx"

# Función para cargar el archivo certificado
def cargar_certificado(file):
    # Cargar los datos desde la fila 27 hasta la 127, considerando la fila 26 como cabecera
    df = pd.read_excel(file, sheet_name=0, header=25, skiprows=0, nrows=101)
    
    # Verificar las columnas del DataFrame cargado
    st.write("Columnas del archivo certificado:")
    st.write(df.columns)  # Mostrar las columnas para verificar su nombre
    st.write(df.head())  # Mostrar las primeras filas para verificar los datos
    return df

# Función para cargar la plantilla desde la ruta local
def cargar_plantilla():
    try:
        plantilla = openpyxl.load_workbook(PLANTILLA_PATH)
        return plantilla
    except FileNotFoundError:
        st.error(f"El archivo {PLANTILLA_PATH} no se encuentra en el directorio.")
        return None

# Función para aplicar filtros y copiar los datos filtrados a la plantilla
def procesar_datos(certificado, plantilla):
    # Filtrar y preparar los datos para las hojas correspondientes
    hoja_o = plantilla["O"]
    hoja_dp = plantilla["DP"]
    hoja_std = plantilla["STD"]
    
    # Rellenar los valores vacíos con la palabra "HOLA"
    certificado.fillna("HOLA", inplace=True)

    # Verificar los nombres de las columnas
    st.write("Columnas del DataFrame procesado:")
    st.write(certificado.columns)

    # Filtrar para la hoja "O" (sin "DEND" ni "DSTD") usando los índices
    datos_o = certificado[~certificado.iloc[:, 14].str.contains('DEND|DSTD', na=False)]  # Columna 'O' es el índice 14
    for i, row in datos_o.iterrows():
        hoja_o.append([row.iloc[0], row.iloc[1], row.iloc[2], row.iloc[3], row.iloc[4], row.iloc[5], row.iloc[6], row.iloc[7],
                       row.iloc[8], row.iloc[9], row.iloc[10], row.iloc[11], row.iloc[13], None, row.iloc[16], None, row.iloc[17]])

    # Filtrar para la hoja "DP" (con "DEND")
    datos_dp = certificado[certificado.iloc[:, 14].str.contains('DEND', na=False)]  # Columna 'O' es el índice 14
    for i, row in datos_dp.iterrows():
        hoja_dp.append([row.iloc[0], row.iloc[1], row.iloc[2], row.iloc[3], row.iloc[4], row.iloc[5], row.iloc[6], row.iloc[7],
                       row.iloc[8], row.iloc[9], row.iloc[10], row.iloc[14], row.iloc[13], row.iloc[16], None, row.iloc[17]])

    # Filtrar para la hoja "STD" (con "DSTD")
    datos_std = certificado[certificado.iloc[:, 14].str.contains('DSTD', na=False)]  # Columna 'O' es el índice 14
    for i, row in datos_std.iterrows():
        hoja_std.append([row.iloc[0], row.iloc[4], row.iloc[6], row.iloc[7], row.iloc[8], row.iloc[9], row.iloc[10], row.iloc[16],
                         row.iloc[13], row.iloc[16], None, row.iloc[17]])

    # Guardar la plantilla con los nuevos datos
    plantilla.save("plantilla_export_modificada.xlsx")

# Función para generar y permitir la descarga de una hoja a CSV
def generar_csv(hoja_nombre, plantilla):
    hoja = plantilla[hoja_nombre]
    
    # Convertir la hoja en un DataFrame
    df = pd.DataFrame(hoja.values)
    
    # Eliminar las palabras "HOLA" de los datos antes de exportar
    df = df.replace("HOLA", "")
    
    # Convertir el DataFrame a un archivo CSV en memoria
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    
    return csv_buffer

# Interfaz de usuario en Streamlit
st.title("Procesar Certificado y Exportar Datos")

# Subir archivo certificado
uploaded_file = st.file_uploader("Sube el archivo certificado.xlsx", type=["xlsx"])

if uploaded_file is not None:
    # Cargar y procesar el archivo certificado
    certificado = cargar_certificado(uploaded_file)
    plantilla = cargar_plantilla()

    if plantilla:
        # Procesar datos y copiarlos a la plantilla
        procesar_datos(certificado, plantilla)
    
        st.success("Los datos han sido procesados y copiados exitosamente.")
    
        # Generar botones para exportar a CSV
        if st.button('Exportar hoja O a CSV'):
            csv_o = generar_csv('O', plantilla)
            st.download_button('Descargar CSV de hoja O', csv_o, file_name="hoja_O.csv", mime='text/csv')
    
        if st.button('Exportar hoja DP a CSV'):
            csv_dp = generar_csv('DP', plantilla)
            st.download_button('Descargar CSV de hoja DP', csv_dp, file_name="hoja_DP.csv", mime='text/csv')
    
        if st.button('Exportar hoja STD a CSV'):
            csv_std = generar_csv('STD', plantilla)
            st.download_button('Descargar CSV de hoja STD', csv_std, file_name="hoja_STD.csv", mime='text/csv')
