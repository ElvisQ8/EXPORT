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
                    0: "A",  # Columna 0 del DataFrame a la columna A de la hoja "O"
                    1: "B",  # Columna 1 del DataFrame a la columna B de la hoja "O"
                    2: "C",  # Columna 2 del DataFrame a la columna C de la hoja "O"
                    3: "D",  # Columna 3 del DataFrame a la columna D de la hoja "O"
                    4: "E",  # Columna 4 del DataFrame a la columna E de la hoja "O"
                    5: "F",  # Columna 5 del DataFrame a la columna F de la hoja "O"
                    6: "G",  # Columna 6 del DataFrame a la columna G de la hoja "O"
                    7: "H",  # Columna 7 del DataFrame a la columna H de la hoja "O"
                    8: "I",  # Columna 8 del DataFrame a la columna I de la hoja "O"
                    9: "J",  # Columna 9 del DataFrame a la columna J de la hoja "O"
                    10: "K",  # Columna 10 del DataFrame a la columna K de la hoja "O"
                    11: "L",  # Columna 11 del DataFrame a la columna L de la hoja "O"
                    13: "M",  # Columna 13 del DataFrame a la columna M de la hoja "O"
                    16: "O",  # Columna 16 del DataFrame a la columna O de la hoja "O"
                    17: "Q"   # Columna 17 del DataFrame a la columna Q de la hoja "O"
                },
                "DP": {
                    0: "A",  # Columna 0 del DataFrame a la columna A de la hoja "DP"
                    1: "B",  # Columna 1 del DataFrame a la columna B de la hoja "DP"
                    2: "C",  # Columna 2 del DataFrame a la columna C de la hoja "DP"
                    3: "D",  # Columna 3 del DataFrame a la columna D de la hoja "DP"
                    4: "E",  # Columna 4 del DataFrame a la columna E de la hoja "DP"
                    6: "F",  # Columna 6 del DataFrame a la columna F de la hoja "DP"
                    7: "G",  # Columna 7 del DataFrame a la columna G de la hoja "DP"
                    8: "H",  # Columna 8 del DataFrame a la columna H de la hoja "DP"
                    9: "I",  # Columna 9 del DataFrame a la columna I de la hoja "DP"
                    10: "J",  # Columna 10 del DataFrame a la columna J de la hoja "DP"
                    14: "K",  # Columna 14 del DataFrame a la columna K de la hoja "DP"
                    13: "L",  # Columna 13 del DataFrame a la columna L de la hoja "DP"
                    16: "M",  # Columna 16 del DataFrame a la columna M de la hoja "DP"
                    17: "O"   # Columna 17 del DataFrame a la columna O de la hoja "DP"
                },
                "STD": {
                    0: "A",  # Columna 0 del DataFrame a la columna A de la hoja "STD"
                    4: "B",  # Columna 4 del DataFrame a la columna B de la hoja "STD"
                    6: "C",  # Columna 6 del DataFrame a la columna C de la hoja "STD"
                    7: "D",  # Columna 7 del DataFrame a la columna D de la hoja "STD"
                    8: "E",  # Columna 8 del DataFrame a la columna E de la hoja "STD"
                    9: "F",  # Columna 9 del DataFrame a la columna F de la hoja "STD"
                    10: "G", # Columna 10 del DataFrame a la columna G de la hoja "STD"
                    "PECLSTDEN02": "H",  # Columna "PECLSTDEN02" en el DataFrame a la columna H de la hoja "STD"
                    13: "I",  # Columna 13 del DataFrame a la columna I de la hoja "STD"
                    16: "J",  # Columna 16 del DataFrame a la columna J de la hoja "STD"
                    17: "L"   # Columna 17 del DataFrame a la columna L de la hoja "STD"
                }
            }

            # Copiar los datos de acuerdo con el mapeo de columnas
            for i, col in enumerate(df_filtered.columns):
                if i in column_mapping[sheet_name]:
                    # Obtener la columna de la plantilla que corresponde a la columna del DataFrame
                    col_letter = column_mapping[sheet_name][i]
                    # Escribir los datos de la columna en la columna correspondiente de la hoja de plantilla
                    writer.sheets[sheet_name].write_column(col_letter + '2', df_filtered[col].values)

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
