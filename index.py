import pandas as pd
import streamlit as st
from io import StringIO

# Cargar el archivo certificado.xlsx
def load_data(file_path):
    return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=26, usecols="A:R", nrows=101)

# Función para eliminar "hola" y la fila 2 de la hoja "O"
def clean_data(df, sheet_name):
    df_cleaned = df[df != 'hola']
    if sheet_name == "O":
        df_cleaned = df_cleaned.drop(index=1)  # Eliminar la fila 2 de la hoja "O"
    return df_cleaned

# Función para copiar los datos según el mapeo y modificar la plantilla
def copy_data_to_template(df, sheet_name, selected_name, template_file):
    # Cargar la plantilla existente
    template = pd.ExcelFile(template_file)

    with StringIO() as output:
        # Filtrar y copiar los datos en la hoja correspondiente
        if sheet_name == "O":
            df_filtered = df[~df[14].str.contains('DSTD|DEND', na=False)]
            df_filtered[13] = selected_name  # Colocar el nombre en la columna "N" de la hoja "O"
        elif sheet_name == "DP":
            df_filtered = df[df[14].str.contains('DEND', na=False)]
            df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "DP"
        elif sheet_name == "STD":
            df_filtered = df[df[14].str.contains('DSTD', na=False)]
            df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "STD"
        
        # Escribir los datos en formato CSV sin incluir índices ni cabeceras adicionales
        df_filtered.to_csv(output, index=False, header=True)

        output.seek(0)  # Asegurarse de que el flujo esté al principio
        return output.getvalue()

# Crear la interfaz de usuario
st.title("Exportar Datos a Plantilla CSV")
st.write("Selecciona el nombre a agregar a las columnas 'N' (hoja 'O') y 'K' (hoja 'STD'):")

# Selección del nombre
names = ["nombre1", "nombre2", "nombre3"]
selected_name = st.selectbox("Selecciona un nombre", names)

# Cargar el archivo Excel subido por el usuario
uploaded_file = st.file_uploader("Sube el archivo .xlsx", type=["xlsx"])

# Cargar la plantilla
template_file = "plantilla_export.xlsx"  # Asegúrate de que esta plantilla esté disponible en tu entorno

if uploaded_file is not None:
    # Cargar los datos
    df = load_data(uploaded_file)

    # Botón para exportar la hoja "O"
    if st.button('Exportar Hoja O'):
        df_cleaned = clean_data(df, "O")
        file_o = copy_data_to_template(df_cleaned, "O", selected_name, template_file)
        st.download_button("Descargar Hoja O como CSV", data=file_o, file_name="plantilla_O.csv")

    # Botón para exportar la hoja "DP"
    if st.button('Exportar Hoja DP'):
        df_clea
