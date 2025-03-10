import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title='Procesador de Excel', layout='centered')
st.title('Carga y Procesamiento de Datos en Excel')

# Cargar el archivo Excel desde el usuario
uploaded_file = st.file_uploader("Sube tu archivo .xlsx", type=["xlsx"])

if uploaded_file:
    # Leer el archivo Excel
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=26)  # Desde fila 27 en adelante
    st.write("Vista previa de los datos cargados:")
    st.dataframe(df)
    
    # Filtrado de datos
    df_O = df[~df.iloc[:, 14].astype(str).str.contains("DEND|DSTD", na=False)][["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "Q", "R"]]
    df_DP = df[df.iloc[:, 14].astype(str).str.contains("DEND", na=False)][["A", "B", "C", "D", "E", "G", "H", "I", "J", "K", "O", "N", "Q", "R"]]
    df_STD = df[df.iloc[:, 14].astype(str).str.contains("DSTD", na=False)][["A", "E", "G", "H", "I", "J", "K", "P", "N", "Q", "R"]]
    
    # Función para exportar los datos como CSV
    def export_csv(dataframe, filename):
        output = BytesIO()
        dataframe.to_csv(output, index=False, encoding='utf-8')
        output.seek(0)
        return output
    
    # Botones para descargar cada archivo filtrado
    st.write("Descarga de archivos procesados:")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button("Descargar O.csv", data=export_csv(df_O, "O.csv"), file_name="O.csv", mime="text/csv")
    with col2:
        st.download_button("Descargar DP.csv", data=export_csv(df_DP, "DP.csv"), file_name="DP.csv", mime="text/csv")
    with col3:
        st.download_button("Descargar STD.csv", data=export_csv(df_STD, "STD.csv"), file_name="STD.csv", mime="text/csv")

st.write("Sube tu archivo y procesa los datos automáticamente.")
