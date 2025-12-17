import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Filtro de Clientes > 30K",
    layout="wide"
)

st.title("📊 Filtro de Clientes con Montos Mayores a 30K")
st.write("Carga uno o varios archivos Excel para identificar clientes con montos elevados.")

# Subida de archivos
uploaded_files = st.file_uploader(
    "📂 Sube uno o más archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file)

            # Columnas por posición (B, C, D, I, M)
            columnas_interes = df.iloc[:, [1, 2, 3, 8, 12]].copy()
            columnas_interes.columns = [
                "DOCUMENTO",
                "NUMERO DE DOCUMENTO",
                "NOMBRE",
                "REFERENCIA",
                "MONTO"
            ]

            # Asegurar que MONTO sea numérico
            columnas_interes["MONTO"] = pd.to_numeric(
                columnas_interes["MONTO"], errors="coerce"
            )

            # Filtrar montos mayores a 30k
            filtrado = columnas_interes[columnas_interes["MONTO"] > 30000]

            # Agregar nombre del archivo (trazabilidad)
            filtrado["Archivo_Origen"] = file.name

            dataframes.append(filtrado)

        except Exception as e:
            st.error(f"Error procesando el archivo {file.name}: {e}")

    if dataframes:
        resultado_final = pd.concat(dataframes, ignore_index=True)

        st.success(f"✅ Se encontraron {len(resultado_final)} clientes con montos > 30K")
        st.dataframe(resultado_final, use_container_width=True)

        # Descargar resultado en Excel
        buffer = BytesIO()
        resultado_final.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        st.download_button(
            label="⬇️ Descargar resultado en Excel",
            data=buffer,
            file_name="clientes_mayores_30k.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("No se encontraron registros con montos mayores a 30K.")
