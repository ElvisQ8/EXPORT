import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Cargar el archivo PLANTILLA_EXPORT.xlsx desde GitHub
@st.cache_data
def load_template():
    template_path = "PLANTILA_EXPORT.xlsx"  # Reemplazar con la URL en GitHub si necesario
    return pd.ExcelFile(template_path)

# Función para procesar los datos
def process_file(uploaded_file, template):
    df_certificado = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Filtrar datos según las condiciones
    df_o = df_certificado[~df_certificado.iloc[:, 14].astype(str).isin(["DEND", "DSTD"])]
    df_dp = df_certificado[df_certificado.iloc[:, 14].astype(str) == "DEND"]
    df_std = df_certificado[df_certificado.iloc[:, 14].astype(str) == "DSTD"]
    
    # Cargar la plantilla
    wb = openpyxl.load_workbook(template)
    
    # Procesar hoja "O"
    sheet_o = wb["O"]
    data_o = df_o.iloc[:, [0,1,2,3,4,5,6,7,8,9,10,11,12,15,16]].values.tolist()
    for i, row in enumerate(data_o, start=27):
        for j, value in enumerate(row, start=1):
            sheet_o.cell(row=i, column=j, value=value)
    
    # Procesar hoja "DP"
    sheet_dp = wb["DP"]
    data_dp = df_dp.iloc[:, [0,1,2,3,4,6,7,8,9,10,13,12,15,16]].values.tolist()
    for i, row in enumerate(data_dp, start=27):
        for j, value in enumerate(row, start=1):
            sheet_dp.cell(row=i, column=j, value=value)
    
    # Procesar hoja "STD"
    sheet_std = wb["STD"]
    data_std = df_std.iloc[:, [0,4,6,7,8,9,10,13,12,15,16]].values.tolist()
    for i, row in enumerate(data_std, start=27):
        for j, value in enumerate(row, start=1):
            sheet_std.cell(row=i, column=j, value=value)
    
    return wb

# Función para descargar una hoja
def download_sheet(workbook, sheet_name):
    output = BytesIO()
    wb_copy = openpyxl.Workbook()
    sheet_copy = wb_copy.active
    sheet_copy.title = sheet_name
    source_sheet = workbook[sheet_name]
    
    for row in source_sheet.iter_rows():
        sheet_copy.append([cell.value for cell in row])
    
    wb_copy.save(output)
    output.seek(0)
    return output

# Interfaz Streamlit
st.title("Procesamiento de Certificados")
uploaded_file = st.file_uploader("Sube el archivo certificado.xlsx", type=["xlsx"])

template = load_template()
if uploaded_file:
    wb_processed = process_file(uploaded_file, "PLANTILA_EXPORT.xlsx")
    
    st.success("Datos procesados correctamente.")
    
    for sheet in ["O", "DP", "STD"]:
        output = download_sheet(wb_processed, sheet)
        st.download_button(
            label=f"Descargar hoja {sheet}",
            data=output,
            file_name=f"{sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

