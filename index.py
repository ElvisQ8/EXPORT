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
    # Leer el archivo Excel tomando la fila 25 como encabezado
    df = pd.read_excel(uploaded_file, sheet_name=0, header=25)
    st.write("Vista previa de los datos cargados:")
    st.dataframe(df)
    
    # Mostrar nombres de columnas para depuración
    st.write("Nombres de columnas detectados en el archivo:")
    st.write(df.columns.tolist())
    
    # Verificar si las columnas existen antes de seleccionarlas
    expected_columns_O = ["Holeid", "From", "To", "Sample number", "Displaced volume (g)", "Wet weight (g)", 
                          "Dry weight (g)", "Coated dry weight (g)", "Weight in water (g)", "Coated weight in water (g)", 
                          "Coat density", "moisture", "Determination method", "Date", "comments"]
    expected_columns_DP = ["hole_number", "depth_from", "depth_to", "sample", "displaced_volume_g_D", "dry_weight_g_D", 
                           "coated_dry_weight_g_D", "weight_water_g", "coated_wght_water_g", "coat_density", "QC_type", 
                           "determination_method", "density_date", "comments"]
    expected_columns_STD = ["hole_number", "displaced_volume_g", "dry_weight_g", "coated_dry_weight_g", "weight_water_g", 
                            "coated_wght_water_g", "coat_density", "DSTD_id", "determination_method", "density_date", "comments"]
    
    available_columns = df.columns.tolist()
    df_O = df[~df.iloc[:, 14].astype(str).str.contains("DEND|DSTD", na=False)][[col for col in expected_columns_O if col in available_columns]]
    df_DP = df[df.iloc[:, 14].astype(str).str.contains("DEND", na=False)][[col for col in expected_columns_DP if col in available_columns]]
    df_STD = df[df.iloc[:, 14].astype(str).str.contains("DSTD", na=False)][[col for col in expected_columns_STD if col in available_columns]]
    
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
