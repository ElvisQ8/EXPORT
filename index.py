import streamlit as st
import pandas as pd
import io

# Cargar el archivo PLANTILLA_EXPORT.xlsx desde el almacenamiento local
PLANTILLA_PATH = "PLANTILA_EXPORT.xlsx"

# Función para filtrar y extraer datos
def filtrar_y_copiar_datos(certificado_df, hoja):
    if hoja == "O":
        df_filtrado = certificado_df[~certificado_df.iloc[:, 14].astype(str).str.contains("DEND|DSTD", na=False)]
        columnas = ["Holeid", "From", "To", "Sample number", "Displaced volume (g)", "Wet weight (g)", "Dry weight (g)",
                    "Coated dry weight (g)", "Weight in water (g)", "Coated weight in water (g)", "Coat density",
                    "moisture", "Determination method", "Laboratory", "Date", "Responsible", "comments"]
        df_final = df_filtrado.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, -1, 16, -1, 17]]
    elif hoja == "DP":
        df_filtrado = certificado_df[certificado_df.iloc[:, 14].astype(str).str.contains("DEND", na=False)]
        columnas = ["hole_number", "depth_from", "depth_to", "sample", "displaced_volume_g_D", "dry_weight_g_D",
                    "coated_dry_weight_g_D", "weight_water_g", "coated_wght_water_g", "coat_density", "QC_type",
                    "determination_method", "density_date", "responsible", "comments"]
        df_final = df_filtrado.iloc[:, [0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 14, 13, 16, -1, 17]]
    elif hoja == "STD":
        df_filtrado = certificado_df[certificado_df.iloc[:, 14].astype(str).str.contains("DSTD", na=False)]
        columnas = ["hole_number", "displaced_volume_g", "dry_weight_g", "coated_dry_weight_g", "weight_water_g",
                    "coated_wght_water_g", "coat_density", "DSTD_id", "determination_method", "density_date", 
                    "responsable", "comments"]
        df_final = df_filtrado.iloc[:, [0, 4, 6, 7, 8, 9, 10, 15, 13, 16, -1, 17]]
    else:
        return None
    
    df_final.columns = columnas
    return df_final

st.title("Procesador de Certificados")

certificado_file = st.file_uploader("Sube el archivo certificado.xlsx", type=["xlsx"])

if certificado_file:
    certificado_df = pd.read_excel(certificado_file, sheet_name=None)
    hoja_certificado = list(certificado_df.keys())[0]  # Tomar la primera hoja
    certificado_df = certificado_df[hoja_certificado].iloc[27:]
    
    plantilla = pd.ExcelFile(PLANTILLA_PATH)
    
    hojas = ["O", "DP", "STD"]
    
    for hoja in hojas:
        df_filtrado = filtrar_y_copiar_datos(certificado_df, hoja)
        
        if df_filtrado is not None:
            st.write(f"Vista previa de la hoja {hoja}:")
            st.dataframe(df_filtrado)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_filtrado.to_excel(writer, sheet_name=hoja, index=False)
                writer.save()
            
            st.download_button(
                label=f"Descargar {hoja}.xlsx",
                data=buffer.getvalue(),
                file_name=f"{hoja}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
