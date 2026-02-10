import streamlit as st
import pandas as pd
from io import BytesIO

from utils import normalize_columns, show_table_es_grouped
from rules import load_rules_csv
from processors import process_by_rules

st.set_page_config(page_title="LIQUIDACIONES (por reglas CSV)", page_icon="🏦", layout="wide")

st.sidebar.header("Parámetros")
use_rules = st.sidebar.checkbox("Usar reglas CSV (reglas_apartamentos.csv)", value=True)
rules_path = st.sidebar.text_input("Ruta CSV reglas", value="reglas_apartamentos.csv")
uploaded = st.file_uploader("Sube el archivo de reservas (.xlsx / .csv)", type=["xlsx", "csv"])

if uploaded is None:
    st.info("Sube un archivo de reservas para empezar.")
else:
    try:
        if str(uploaded.name).lower().endswith(".csv"):
            df_raw = pd.read_csv(uploaded)
        else:
            df_raw = pd.read_excel(uploaded, engine="openpyxl")
    except Exception as e:
        st.error(f"Error leyendo el fichero: {e}")
        st.stop()

    df_norm = normalize_columns(df_raw)
    st.markdown(f"Filas detectadas: **{len(df_norm)}**")
    if st.checkbox("Mostrar primeras filas", value=False):
        st.dataframe(df_norm.head(10))

    rules = load_rules_csv(rules_path) if use_rules else {}
    if use_rules:
        st.sidebar.markdown(f"Reglas cargadas: **{len(rules)}**")

    if st.button("Generar liquidación"):
        try:
            out, warn = process_by_rules(df_norm, rules)
            if warn:
                st.warning(f"Reservas con portal vacío y comisión>0: {warn}")
            show_table_es_grouped(out, "Tabla de liquidaciones (por reglas)", group_col="Alojamiento")

            # preparar descarga Excel
            b = BytesIO()
            with pd.ExcelWriter(b, engine="openpyxl") as writer:
                out.to_excel(writer, sheet_name="Liquidaciones", index=False)
            b.seek(0)
            st.download_button("Descargar Excel", data=b, file_name="liquidaciones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error generando liquidación: {e}")
            raise