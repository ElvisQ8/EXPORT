import pandas as pd
import streamlit as st
from io import BytesIO

# Cargar el archivo certificado.xlsx
def load_data(file_path):
    return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=27, usecols="A:R", nrows=101)

# Función para eliminar "hola" de las columnas
def clean_data(df, sheet_name):
    df_cleaned = df[df != 'hola']
    return df_cleaned

# Función para copiar los datos según el mapeo y modificar la plantilla
def copy_data_to_template(df, sheet_name, selected_name, template_file):
    # Cargar la plantilla existente
    template = pd.ExcelFile(template_file)

    with BytesIO() as output:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Escribir solo la hoja seleccionada de la plantilla (sin sobrescribir la primera fila)
            if sheet_name in template.sheet_names:
                temp_df = template.parse(sheet_name)
                temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

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
            
            # Definir el mapeo de columnas que vamos a copiar de los datos a las columnas de la plantilla
            column_mapping = {
                "O": {
                    0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F", 6: "G", 7: "H",
                    8: "I", 9: "J", 10: "K", 11: "L", 13: "M", 16: "O", 17: "Q"
                },
                "DP": {
                    0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 6: "F", 7: "G", 8: "H",
                    9: "I", 10: "J", 14: "K", 13: "L", 16: "M", 17: "O"
                },
                "STD": {
                    0: "A", 4: "B", 6: "C", 7: "D", 8: "E", 9: "F", 10: "G", "PECLSTDEN02": "H",
                    13: "I", 16: "J", 17: "L"
                }
            }

            # Copiar los datos de acuerdo con el mapeo de columnas
            for i, col in enumerate(df_filtered.columns):
                if i in column_mapping[sheet_name]:
                    # Obtener la columna de la plantilla que corresponde a la columna del DataFrame
                    col_letter = column_mapping[sheet_name][i]
                    # Escribir los datos de la columna en la columna correspondiente de la hoja de plantilla
                    col_data = df_filtered[col].values
                    # Convertir las columnas a las celdas correspondientes en el archivo Excel
                    for row_idx, value in enumerate(col_data, start=2):  # Comenzamos desde la fila 2
                        writer.sheets[sheet_name].write(f'{col_letter}{row_idx}', value)

        output.seek(0)  # Asegurarse de que el flujo esté al principio
        return output.getvalue()

# Crear la interfaz de usuario
st.title("Exportar Datos a Plantilla Excel")
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
        st.download_button("Descargar Hoja O", data=file_o, file_name="plantilla_O.xlsx")

    # Botón para exportar la hoja "DP"
    if st.button('Exportar Hoja DP'):
        df_cleaned = clean_data(df, "DP")
        file_dp = copy_data_to_template(df_cleaned, "DP", selected_name, template_file)
        st.download_button("Descargar Hoja DP", data=file_dp, file_name="plantilla_DP.xlsx")

    # Botón para exportar la hoja "STD"
    if st.button('Exportar Hoja STD'):
        df_cleaned = clean_data(df, "STD")
        file_std = copy_data_to_template(df_cleaned, "STD", selected_name, template_file)
        st.download_button("Descargar Hoja STD", data=file_std, file_name="plantilla_STD.xlsx")
