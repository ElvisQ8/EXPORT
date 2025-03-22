import pandas as pd
import streamlit as st
from io import BytesIO

# Cargar el archivo certificado.xlsx
def load_data(file_path):
    return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=27, usecols="A:R", nrows=101)

# Función para limpiar los datos eliminando "hola"
def clean_data(df):
    # Eliminar las celdas con la palabra 'hola'
    df_cleaned = df[df != 'hola']
    return df_cleaned

# Función para copiar los datos a la plantilla y escribirlos en las hojas correspondientes
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
                # Mapeo de columnas para "O"
                columns_mapping_o = [
                    (0, 'B'), (1, 'C'), (2, 'D'), (3, 'E'), (4, 'F'),
                    (5, 'G'), (6, 'H'), (7, 'I'), (8, 'J'), (9, 'K'),
                    (10, 'L'), (11, 'M'), (13, 'O'), (16, 'Q')
                ]
                for df_col, template_col in columns_mapping_o:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

            elif sheet_name == "DP":
                df_filtered = df[df[14].str.contains('DEND', na=False)]
                df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "DP"
                # Mapeo de columnas para "DP"
                columns_mapping_dp = [
                    (0, 'A'), (1, 'B'), (2, 'C'), (3, 'D'), (4, 'E'),
                    (6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'),
                    (14, 'K'), (13, 'L'), (16, 'M'), (17, 'O')
                ]
                for df_col, template_col in columns_mapping_dp:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

            elif sheet_name == "STD":
                df_filtered = df[df[14].str.contains('DSTD', na=False)]
                df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "STD"
                # Mapeo de columnas para "STD"
                columns_mapping_std = [
                    (0, 'A'), (4, 'B'), (6, 'C'), (7, 'D'), (8, 'E'),
                    (9, 'F'), (10, 'G'), (11, 'H'), (13, 'I'), (16, 'J'),
                    (17, 'L')
                ]
                # En este caso, el valor "PECLSTDEN02" lo tratamos como un valor específico
                peclstd_value = "PECLSTDEN02"
                writer.sheets[sheet_name].write_column('H2', [peclstd_value] * len(df_filtered))  # Asignamos este valor en la columna H
                for df_col, template_col in columns_mapping_std:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

        # Convertir el archivo a CSV para la descarga
        output.seek(0)  # Asegurarse de que el flujo esté al principio
        df_csv = pd.read_excel(output, sheet_name=sheet_name)  # Leer el archivo modificado en el buffer
        csv_output = BytesIO()
        df_csv.to_csv(csv_output, index=False, sep=';', encoding='utf-8')  # Convertir a CSV
        csv_output.seek(0)
        return csv_output.getvalue()

# Crear la interfaz de usuario en Streamlit
st.title("Exportar Datos a Plantilla Excel")
st.write("Selecciona el nombre a agregar a las columnas 'N' (hoja 'O') y 'K' (hoja 'STD'): ")

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
        df_cleaned = clean_data(df)
        file_o = copy_data_to_template(df_cleaned, "O", selected_name, template_file)
        st.download_button("Descargar Hoja O como CSV", data=file_o, file_name="plantilla_O.csv")

    # Botón para exportar la hoja "DP"
    if st.button('Exportar Hoja DP'):
        df_cleaned = clean_data(df)
        file_dp = copy_data_to_template(df_cleaned, "DP", selected_name, template_file)
        st.download_button("Descargar Hoja DP como CSV", data=file_dp, file_name="plantilla_DP.csv")

    # Botón para exportar la hoja "STD"
    if st.button('Exportar Hoja STD'):
        df_cleaned = clean_data(df)
        file_std = copy_data_to_template(df_cleaned, "STD", selected_name, template_file)
        st.download_button("Descargar Hoja STD como CSV", data=file_std, file_name="plantilla_STD.csv")
