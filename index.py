import pandas as pd
import streamlit as st
from io import BytesIO

# Cargar el archivo certificado.xlsx
def load_data(file_path):
    return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=27, usecols="A:R", nrows=101)

# Función para eliminar "hola" y la fila 2 de la hoja "O"
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

            # Mapeo de columnas para la hoja "O"
            if sheet_name == "O":
                # Mapeo columna por columna de acuerdo con las instrucciones
                writer.sheets[sheet_name].write_column('B2', df_filtered[0].values)  # Columna 0 -> Columna B
                writer.sheets[sheet_name].write_column('C2', df_filtered[1].values)  # Columna 1 -> Columna C
                writer.sheets[sheet_name].write_column('D2', df_filtered[2].values)  # Columna 2 -> Columna D
                writer.sheets[sheet_name].write_column('E2', df_filtered[3].values)  # Columna 3 -> Columna E
                writer.sheets[sheet_name].write_column('F2', df_filtered[4].values)  # Columna 4 -> Columna F
                writer.sheets[sheet_name].write_column('G2', df_filtered[5].values)  # Columna 5 -> Columna G
                writer.sheets[sheet_name].write_column('H2', df_filtered[6].values)  # Columna 6 -> Columna H
                writer.sheets[sheet_name].write_column('I2', df_filtered[7].values)  # Columna 7 -> Columna I
                writer.sheets[sheet_name].write_column('J2', df_filtered[8].values)  # Columna 8 -> Columna J
                writer.sheets[sheet_name].write_column('K2', df_filtered[9].values)  # Columna 9 -> Columna K
                writer.sheets[sheet_name].write_column('L2', df_filtered[10].values)  # Columna 10 -> Columna L
                writer.sheets[sheet_name].write_column('M2', df_filtered[11].values)  # Columna 11 -> Columna M
                writer.sheets[sheet_name].write_column('O2', df_filtered[13].values)  # Columna 13 -> Columna O
                writer.sheets[sheet_name].write_column('Q2', df_filtered[16].values)  # Columna 16 -> Columna Q

            # Mapeo para la hoja "DP"
            elif sheet_name == "DP":
                writer.sheets[sheet_name].write_column('B2', df_filtered[0].values)  # Columna 0 -> Columna B
                writer.sheets[sheet_name].write_column('C2', df_filtered[1].values)  # Columna 1 -> Columna C
                writer.sheets[sheet_name].write_column('D2', df_filtered[2].values)  # Columna 2 -> Columna D
                writer.sheets[sheet_name].write_column('E2', df_filtered[3].values)  # Columna 3 -> Columna E
                writer.sheets[sheet_name].write_column('F2', df_filtered[6].values)  # Columna 6 -> Columna F
                writer.sheets[sheet_name].write_column('G2', df_filtered[7].values)  # Columna 7 -> Columna G
                writer.sheets[sheet_name].write_column('H2', df_filtered[8].values)  # Columna 8 -> Columna H
                writer.sheets[sheet_name].write_column('I2', df_filtered[9].values)  # Columna 9 -> Columna I
                writer.sheets[sheet_name].write_column('J2', df_filtered[10].values)  # Columna 10 -> Columna J
                writer.sheets[sheet_name].write_column('K2', df_filtered[14].values)  # Columna 14 -> Columna K
                writer.sheets[sheet_name].write_column('L2', df_filtered[13].values)  # Columna 13 -> Columna L
                writer.sheets[sheet_name].write_column('M2', df_filtered[16].values)  # Columna 16 -> Columna M
                writer.sheets[sheet_name].write_column('O2', df_filtered[17].values)  # Columna 17 -> Columna O

            # Mapeo para la hoja "STD"
            elif sheet_name == "STD":
                writer.sheets[sheet_name].write_column('B2', df_filtered[0].values)  # Columna 0 -> Columna B
                writer.sheets[sheet_name].write_column('C2', df_filtered[4].values)  # Columna 4 -> Columna C
                writer.sheets[sheet_name].write_column('D2', df_filtered[6].values)  # Columna 6 -> Columna D
                writer.sheets[sheet_name].write_column('E2', df_filtered[7].values)  # Columna 7 -> Columna E
                writer.sheets[sheet_name].write_column('F2', df_filtered[8].values)  # Columna 8 -> Columna F
                writer.sheets[sheet_name].write_column('G2', df_filtered[9].values)  # Columna 9 -> Columna G
                writer.sheets[sheet_name].write_column('H2', df_filtered[10].values)  # Columna 10 -> Columna H
                writer.sheets[sheet_name].write_column('I2', df_filtered[13].values)  # Columna 13 -> Columna I
                writer.sheets[sheet_name].write_column('J2', df_filtered[16].values)  # Columna 16 -> Columna J
                writer.sheets[sheet_name].write_column('L2', df_filtered[17].values)  # Columna 17 -> Columna L

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
