import streamlit as st
import pandas as pd

# Función para cargar el archivo Excel
def load_excel(file):
    return pd.read_excel(file, sheet_name=None)

# Función para filtrar y copiar los datos en la plantilla
def process_data(certificado_df, plantilla_df):
    # Hoja "O" - Filtrar y copiar datos que no contengan "DEND" ni "DSTD"
    hoja_o = certificado_df[certificado_df['Unnamed: 0'].str.contains('DEND|DSTD', na=False) == False]
    hoja_o = hoja_o[['Holeid', 'From', 'To', 'Sample number', 'Displaced volume (g)', 'Wet weight (g)',
                     'Dry weight (g)', 'Coated dry weight (g)', 'Weight in water (g)', 'Coated weight in water (g)',
                     'Coat density', 'moisture', 'Determination method', 'Laboratory', 'Date', 'Responsible', 'comments']]
    hoja_o.columns = ['Holeid', 'From', 'To', 'Sample number', 'Displaced volume (g)', 'Wet weight (g)', 'Dry weight (g)',
                      'Coated dry weight (g)', 'Weight in water (g)', 'Coated weight in water (g)', 'Coat density',
                      'moisture', 'Determination method', 'Laboratory', 'Date', 'Responsible', 'comments']

    plantilla_df['O'] = hoja_o
    
    # Hoja "DP" - Filtrar y copiar datos que contengan "DEND"
    hoja_dp = certificado_df[certificado_df['Unnamed: 0'].str.contains('DEND', na=False)]
    hoja_dp = hoja_dp[['hole_number', 'depth_from', 'depth_to', 'sample', 'displaced_volume_g_D', 'dry_weight_g_D',
                       'coated_dry_weight_g_D', 'weight_water_g', 'coated_wght_water_g', 'coat_density', 'QC_type',
                       'determination_method', 'density_date', 'responsible', 'comments']]
    hoja_dp.columns = ['hole_number', 'depth_from', 'depth_to', 'sample', 'displaced_volume_g_D', 'dry_weight_g_D',
                       'coated_dry_weight_g_D', 'weight_water_g', 'coated_wght_water_g', 'coat_density', 'QC_type',
                       'determination_method', 'density_date', 'responsible', 'comments']

    plantilla_df['DP'] = hoja_dp
    
    # Hoja "STD" - Filtrar y copiar datos que contengan "DSTD"
    hoja_std = certificado_df[certificado_df['Unnamed: 0'].str.contains('DSTD', na=False)]
    hoja_std = hoja_std[['hole_number', 'displaced_volume_g', 'dry_weight_g', 'coated_dry_weight_g', 'weight_water_g',
                         'coated_wght_water_g', 'coat_density', 'DSTD_id', 'determination_method', 'density_date',
                         'responsible', 'comments']]
    hoja_std.columns = ['hole_number', 'displaced_volume_g', 'dry_weight_g', 'coated_dry_weight_g', 'weight_water_g',
                        'coated_wght_water_g', 'coat_density', 'DSTD_id', 'determination_method', 'density_date',
                        'responsible', 'comments']

    plantilla_df['STD'] = hoja_std
    
    return plantilla_df

# Función para exportar a CSV
def export_to_csv(sheet, filename):
    sheet.to_csv(filename, index=False)

# Interfaz Streamlit
st.title("Procesador de Datos Excel")

# Cargar archivo "certificado.xlsx"
certificado_file = st.file_uploader("Cargar archivo certificado.xlsx", type=["xlsx"])

if certificado_file:
    # Cargar los datos
    certificado_df = pd.read_excel(certificado_file, sheet_name='PECLD07826', skiprows=26)  # Fila 27 es la 26 en 0-indexed
    
    # Cargar la plantilla desde el archivo PLANTILLA_EXPORT.xlsx (se asume que está en el mismo directorio que index.py)
    plantilla_file = "PLANTILA_EXPORT.xlsx"
    plantilla_df = pd.read_excel(plantilla_file, sheet_name=None)  # Cargar todas las hojas de la plantilla

    # Procesar los datos y copiarlos a la plantilla
    plantilla_df = process_data(certificado_df, plantilla_df)

    # Mostrar vista previa de las hojas procesadas
    st.subheader("Vista previa de los datos procesados")
    st.write(plantilla_df['O'].head())  # Muestra las primeras filas de la hoja "O"
    st.write(plantilla_df['DP'].head())  # Muestra las primeras filas de la hoja "DP"
    st.write(plantilla_df['STD'].head())  # Muestra las primeras filas de la hoja "STD"

    # Botones para exportar a CSV
    if st.button('Exportar hoja O a CSV'):
        export_to_csv(plantilla_df['O'], "hoja_O.csv")
        st.success("Hoja 'O' exportada como CSV")

    if st.button('Exportar hoja DP a CSV'):
        export_to_csv(plantilla_df['DP'], "hoja_DP.csv")
        st.success("Hoja 'DP' exportada como CSV")

    if st.button('Exportar hoja STD a CSV'):
        export_to_csv(plantilla_df['STD'], "hoja_STD.csv")
        st.success("Hoja 'STD' exportada como CSV")
