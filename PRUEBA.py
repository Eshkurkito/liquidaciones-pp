import streamlit as st
import pandas as pd
import numpy as np
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="LIQUIDACIONES Dinámicas", page_icon="🏦", layout="wide")

# =========================================================
# UTILIDADES
# =========================================================

def normalize_columns(df):
    rename_map = {
        "Nombre alojamiento": "Alojamiento",
        "Fecha de entrada": "Fecha entrada",
        "Fecha de salida": "Fecha salida",
        "Noches": "Noches ocupadas",
        "Alquiler con tasas": "Ingreso alojamiento",
        "Tarifa limpieza": "Ingreso limpieza",
        "Total reserva con tasas": "Total ingresos",
        "Web origen": "Portal",
        "Comisión Portal/Intermediario: Comisión calculada": "Comisión portal",
    }
    df = df.rename(columns=rename_map)

    numeric_cols = [
        "Ingreso alojamiento",
        "Ingreso limpieza",
        "Total ingresos",
        "Comisión portal"
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if "Alojamiento" in df.columns:
        df["Alojamiento"] = df["Alojamiento"].astype(str).str.strip().str.upper()

    return df


# =========================================================
# MOTOR DINÁMICO
# =========================================================

def process_dynamic(df, reglas, vat_pct=21.0):

    df = df.merge(reglas, on="Alojamiento", how="left")

    if df["modelo"].isna().any():
        st.error("Hay alojamientos sin reglas en el CSV.")
        st.stop()

    df["Comisión portal"] = pd.to_numeric(df["Comisión portal"], errors="coerce").fillna(0)

    # Recalcular comisión con IVA si aplica
    mask_booking = df["Portal"].astype(str).str.lower().str.contains("booking", na=False)

    df.loc[
        (mask_booking) & (df["recalcular_comision_con_iva"] == True),
        "Comisión portal"
    ] *= (1 + vat_pct / 100)

    # IVA alquiler si aplica
    df["IVA del alquiler"] = 0
    mask_iva10 = df["iva_alquiler_tipo"] == 10
    df.loc[mask_iva10, "IVA del alquiler"] = (
        df.loc[mask_iva10, "Ingreso alojamiento"]
        - (df.loc[mask_iva10, "Ingreso alojamiento"] / 1.10)
    )

    # Honorarios
    def calc_honorarios(row):

        base = row["Ingreso alojamiento"]

        if row["modelo"] == 3:
            base = base - row["Comisión portal"]

        if row["modelo"] in [4, 5]:
            base = base - row["IVA del alquiler"] - row["Comisión portal"]

        return base * row["pct_honorarios"] * 1.21

    df["Honorarios Florit"] = df.apply(calc_honorarios, axis=1)

    # Limpieza
    df["Gasto limpieza"] = np.where(
        df["usar_limpieza_excel"] == True,
        df["Ingreso limpieza"],
        df["limpieza_fija"].fillna(0)
    )

    # Amenities
    df["Amenities"] = df["amenities"].fillna(0)

    # Totales
    df["Total Gastos"] = (
        df["Comisión portal"]
        + df["Honorarios Florit"]
        + df["Gasto limpieza"]
        + df["Amenities"]
    )

    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]

    return df


# =========================================================
# EXCEL
# =========================================================

def build_excel(df):

    wb = Workbook()
    ws = wb.active
    ws.title = "Liquidación"

    for col_idx, col in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col).font = Font(bold=True)

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    return bio


# =========================================================
# UI
# =========================================================

st.title("📊 Liquidaciones dinámicas por CSV")

start_date = st.date_input("Desde", value=date(date.today().year, date.today().month, 1))
end_date = st.date_input("Hasta", value=date.today())

file_reservas = st.file_uploader("Sube reservas (.xlsx)", type=["xlsx"])
file_reglas = st.file_uploader("Sube reglas_apartamentos.csv", type=["csv"])

if st.button("Generar liquidación"):

    if not file_reservas or not file_reglas:
        st.error("Sube ambos archivos.")
        st.stop()

    df_res = pd.read_excel(file_reservas)
    df_res = normalize_columns(df_res)

    df_res = df_res[
        (df_res["Fecha entrada"] >= pd.to_datetime(start_date))
        & (df_res["Fecha entrada"] <= pd.to_datetime(end_date))
    ]

    reglas = pd.read_csv(file_reglas)
    reglas["Alojamiento"] = reglas["Alojamiento"].astype(str).str.strip().str.upper()

    df_final = process_dynamic(df_res, reglas)

    st.success("Liquidación generada correctamente.")
    st.dataframe(df_final)

    excel_file = build_excel(df_final)

    st.download_button(
        "Descargar Excel",
        excel_file.getvalue(),
        file_name="Liquidacion_dinamica.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
