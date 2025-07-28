import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Convertidor XLS a XLSX")

uploaded_file = st.file_uploader("Sube tu archivo .xls aquí", type=["xls"])

if uploaded_file is not None:
    try:
        # Leer archivo .xls con xlrd
        df = pd.read_excel(uploaded_file, engine='xlrd')

        # Convertir DataFrame a archivo Excel en memoria (.xlsx)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        data = output.getvalue()

        st.success("Archivo convertido correctamente!")

        # Botón para descargar el archivo .xlsx convertido
        st.download_button(
            label="Descargar archivo convertido (.xlsx)",
            data=data,
            file_name="archivo_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error al convertir el archivo: {e}")