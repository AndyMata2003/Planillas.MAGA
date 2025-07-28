import os
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from docx2pdf import convert
import zipfile
from io import BytesIO
import tempfile
import datetime

# --- CONFIGURACIÓN ---
TEMPLATE_PATH = "PLANILLAS DAPCA v3.docx"

# --- CARGA DE ARCHIVO 1 ---
st.title("Generador de Planillas DAPCA")

archivo_1 = st.file_uploader("Sube el archivo 1 (datos del técnico)", type=["xlsx"])
if archivo_1:
    df_datos = pd.read_excel(archivo_1)
    st.success("Archivo 1 cargado")

    # Simulación de archivo 2 (beneficiarios)
    st.info("Simulando archivo de beneficiarios para el ejemplo...")
    datos_beneficiarios = []
    comunidades = df_datos["Comunidad/ Establecimiento"].dropna().unique()

    for comunidad in comunidades:
        for i in range(1, 25):  # Simula 24 beneficiarios por comunidad
            datos_beneficiarios.append({
                "Referencia": comunidad.lower().strip(),
                "PRIMER NOMBRE": f"Nombre{i}",
                "SEGUNDO NOMBRE": "",
                "TERCER NOMBRE": "",
                "PRIMER APELLIDO": f"Apellido{i}",
                "SEGUNDO APELLIDO": "",
                "APELLIDO CASADA": "",
                "CUI": f"123456789010{i:02d}"
            })

    df_beneficiarios = pd.DataFrame(datos_beneficiarios)

    # Botón para generar planillas
    if st.button("📄 Generar planillas PDF por comunidad"):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_zip = os.path.join(tmpdir, "planillas.zip")
            pdf_paths = []

            for comunidad in comunidades:
                ref = comunidad.strip().lower()
                df_filtrado = df_beneficiarios[df_beneficiarios["Referencia"] == ref]
                nombre_completo = (
                    df_filtrado["PRIMER NOMBRE"].fillna("") + " " +
                    df_filtrado["SEGUNDO NOMBRE"].fillna("") + " " +
                    df_filtrado["TERCER NOMBRE"].fillna("") + " " +
                    df_filtrado["PRIMER APELLIDO"].fillna("") + " " +
                    df_filtrado["SEGUNDO APELLIDO"].fillna("") + " " +
                    df_filtrado["APELLIDO CASADA"].fillna("")
                ).str.replace(r'\s+', ' ', regex=True).str.strip()

                df_filtrado["NOMBRE COMPLETO"] = nombre_completo
                df_final = df_filtrado[["NOMBRE COMPLETO", "CUI"]].reset_index(drop=True)

                # Agrupar en bloques de 10
                bloques = [df_final[i:i + 10] for i in range(0, len(df_final), 10)]
                paginas = []

                for bloque in bloques:
                    fila_data = []
                    for idx in range(10):
                        if idx < len(bloque):
                            fila_data.append({
                                "no": idx + 1,
                                "nombre": bloque.iloc[idx]["NOMBRE COMPLETO"],
                                "cui": bloque.iloc[idx]["CUI"]
                            })
                        else:
                            fila_data.append({
                                "no": idx + 1,
                                "nombre": "",
                                "cui": ""
                            })
                    paginas.append({"filas": fila_data})

                # Cargar y renderizar plantilla
                doc = DocxTemplate(TEMPLATE_PATH)

                tecnico_data = df_datos[df_datos["Comunidad/ Establecimiento"].str.lower().str.strip() == ref].iloc[0]

                context = {
                    "codigo": f"DAPA-017-2025",
                    "fecha": datetime.datetime.now().strftime("%d/%m/%Y"),
                    "comunidad": comunidad.title(),
                    "departamento": tecnico_data["Departamento"],
                    "municipio": tecnico_data["Municipio"],
                    "tecnico": tecnico_data["Nombre Técnico"],
                    "dpi": str(tecnico_data["DPI"]),
                    "paginas": paginas
                }

                output_docx = os.path.join(tmpdir, f"{comunidad}.docx")
                output_pdf = os.path.join(tmpdir, f"{comunidad}.pdf")

                doc.render(context)
                doc.save(output_docx)

                convert(output_docx, output_pdf)
                pdf_paths.append(output_pdf)

            # Crear ZIP con todos los PDFs
            with zipfile.ZipFile(output_zip, "w") as zf:
                for pdf in pdf_paths:
                    zf.write(pdf, arcname=os.path.basename(pdf))

            # Descargar ZIP
            with open(output_zip, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar planillas en ZIP",
                    data=f,
                    file_name="planillas_comunidades.zip",
                    mime="application/zip"
                )