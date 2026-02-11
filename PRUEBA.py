import streamlit as st
import pandas as pd
import numpy as np
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font

st.set_page_config(page_title="Liquidaciones Dinámicas", page_icon="🏦", layout="wide")

# =====================================================
# CARGA REGLAS DESDE REPOSITORIO
# =====================================================

@st.cache_data
def load_reglas():
    try:
        reglas = pd.read_csv("reglas_apartamentos.csv", sep=";")
    except FileNotFoundError:
        st.error("No se encuentra reglas_apartamentos.csv en el repositorio.")
        st.stop()

    reglas.columns = reglas.columns.str.strip()

    if "Property" not in reglas.columns:
        st.error(f"Columnas detectadas: {list(reglas.columns)}")
        st.error("No existe columna 'Property' en reglas_apartamentos.csv")
        st.stop()

    reglas["Property"] = (
        reglas["Property"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Convertir booleanos correctamente
    bool_cols = [
        "honorarios_apply_vat",
        "compute_iva_alquiler",
        "treat_empty_portal_as_booking",
        "skip_booking_vat",
        "hon_base_exclude_commission"
    ]

    for col in bool_cols:
        if col in reglas.columns:
            reglas[col] = reglas[col].astype(str).str.upper().map(
                {"TRUE": True, "FALSE": False}
            )

    return reglas


# =====================================================
# NORMALIZACIÓN RESERVAS
# =====================================================

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
        df["Alojamiento"] = (
            df["Alojamiento"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

    if "Fecha entrada" in df.columns:
        df["Fecha entrada"] = pd.to_datetime(df["Fecha entrada"], errors="coerce")

    return df


# =====================================================
# MOTOR DINÁMICO
# =====================================================

def process_dynamic(df, reglas):

    df = df.merge(
        reglas,
        left_on="Alojamiento",
        right_on="Property",
        how="left"
    )

    if df["Property"].isna().any():
        faltantes = df[df["Property"].isna()]["Alojamiento"].unique()
        st.error(f"Alojamientos sin reglas definidas: {faltantes}")
        st.stop()

    # ------------------------------
    # Comisión portal
    # ------------------------------

    df["Comisión portal"] = pd.to_numeric(
        df["Comisión portal"],
        errors="coerce"
    ).fillna(0)

    mask_booking = df["Portal"].astype(str).str.lower().str.contains("booking", na=False)

    df.loc[
        (mask_booking) & (df["skip_booking_vat"] == False),
        "Comisión portal"
    ] *= (1 + df["commission_vat_pct"] / 100)

    # ------------------------------
    # IVA alquiler
    # ------------------------------

    df["IVA del alquiler"] = 0.0

    mask_iva = df["compute_iva_alquiler"] == True

    df.loc[mask_iva, "IVA del alquiler"] = (
        df.loc[mask_iva, "Ingreso alojamiento"]
        - (df.loc[mask_iva, "Ingreso alojamiento"] / 1.10)
    )

    # ------------------------------
    # Honorarios
    # ------------------------------

    def calc_honorarios(row):

        base = row["Ingreso alojamiento"]

        if row["hon_base_exclude_commission"]:
            base -= row["Comisión portal"]

        if row["compute_iva_alquiler"]:
            base -= row["IVA del alquiler"]

        honorarios = base * row["honorarios_pct"]

        if row["honorarios_apply_vat"]:
            honorarios *= (1 + row["honorarios_vat_pct"] / 100)

        return honorarios

    df["Honorarios Florit"] = df.apply(calc_honorarios, axis=1)

    # ------------------------------
    # Limpieza
    # ------------------------------

    df["Gasto limpieza"] = np.where(
        df["cleaning_fee"].notna(),
        df["cleaning_fee"],
        df["Ingreso limpieza"]
    )

    # ------------------------------
    # Amenities
    # ------------------------------

    df["Amenities"] = df["amenities_amount"].fillna(0)

    # ------------------------------
    # Totales
    # ------------------------------

    df["Total Gastos"] = (
        df["Comisión portal"]
        + df["Honorarios Florit"]
        + df["Gasto limpieza"]
        + df["Amenities"]
    )

    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]

    columnas_finales = [
        "Alojamiento",
        "Fecha entrada",
        "Fecha salida",
        "Noches ocupadas",
        "Ingreso alojamiento",
        "IVA del alquiler",
        "Ingreso limpieza",
        "Total ingresos",
        "Portal",
        "Comisión portal",
        "Honorarios Florit",
        "Gasto limpieza",
        "Amenities",
        "Total Gastos",
        "Pago al propietario",
        "Pago recibido"
    ]

    columnas_existentes = [c for c in columnas_finales if c in df.columns]

    return df[columnas_existentes]


# =====================================================
# EXPORTAR EXCEL
# =====================================================

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


# =====================================================
# UI
# =====================================================

st.title("📊 Liquidaciones dinámicas")

start_date = st.date_input("Desde", value=date(date.today().year, date.today().month, 1))
end_date = st.date_input("Hasta", value=date.today())

file_reservas = st.file_uploader("Sube archivo de reservas (.xlsx)", type=["xlsx"])

if st.button("Generar liquidación"):

    if not file_reservas:
        st.error("Sube el archivo de reservas.")
        st.stop()

    reglas = load_reglas()

    df_res = pd.read_excel(file_reservas)
    df_res = normalize_columns(df_res)

    df_res = df_res[
        (df_res["Fecha entrada"] >= pd.to_datetime(start_date))
        & (df_res["Fecha entrada"] <= pd.to_datetime(end_date))
    ]

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
