import streamlit as st
import pandas as pd
import numpy as np
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font

st.set_page_config(page_title="Liquidaciones dinámicas", page_icon="📊", layout="wide")

# =====================================================
# CARGA REGLAS CSV (AUTODETECTA SEPARADOR)
# =====================================================

@st.cache_data
def load_reglas():
    try:
        reglas = pd.read_csv("reglas_apartamentos.csv", sep=None, engine="python")
    except FileNotFoundError:
        st.error("No se encuentra reglas_apartamentos.csv en el repositorio.")
        st.stop()

    reglas.columns = reglas.columns.str.strip()

    if "Property" not in reglas.columns:
        st.error(f"Columnas detectadas en CSV: {list(reglas.columns)}")
        st.error("No existe columna 'Property' en reglas_apartamentos.csv")
        st.stop()

    reglas["Property"] = reglas["Property"].astype(str).str.strip().str.upper()

    bool_cols = [
        "honorarios_apply_vat",
        "compute_iva_alquiler",
        "treat_empty_portal_as_booking",
        "skip_booking_vat",
        "hon_base_exclude_commission"
    ]

    for col in bool_cols:
        if col in reglas.columns:
            reglas[col] = (
                reglas[col]
                .astype(str)
                .str.upper()
                .map({"TRUE": True, "FALSE": False})
                .fillna(False)
            )

    return reglas


# =====================================================
# NORMALIZACIÓN ROBUSTA (COPIADA DEL BACKUP)
# =====================================================

def normalize_columns(df):

    out = df.copy()

    def first_existing(candidates):
        norm_map = {str(c).strip().lower(): c for c in out.columns}
        for cand in candidates:
            key = cand.strip().lower()
            if key in norm_map:
                return norm_map[key]
        return None

    col_aloj = first_existing(["Nombre alojamiento","Alojamiento","Nombre del alojamiento"])
    col_fent = first_existing(["Fecha entrada","Fecha de entrada"])
    col_fsal = first_existing(["Fecha salida","Fecha de salida"])
    col_noch = first_existing(["Noches","Noches ocupadas"])
    col_alq  = first_existing(["Alquiler con tasas","Ingreso alojamiento","Importe alojamiento"])
    col_ext  = first_existing([
        "Ingreso limpieza","Tarifa limpieza","Limpieza","Importe limpieza",
        "Extras con tasas","Gastos de limpieza","Gasto limpieza"
    ])
    col_tot  = first_existing(["Total reserva con tasas","Total ingresos","Total"])
    col_port = first_existing(["Web origen","Portal","Canal"])
    col_comi = first_existing([
        "Comisión Portal/Intermediario: Comisión calculada",
        "Comisión portal","Comisión"
    ])

    rename = {}
    if col_aloj: rename[col_aloj] = "Alojamiento"
    if col_fent: rename[col_fent] = "Fecha entrada"
    if col_fsal: rename[col_fsal] = "Fecha salida"
    if col_noch: rename[col_noch] = "Noches ocupadas"
    if col_alq:  rename[col_alq]  = "Ingreso alojamiento"
    if col_ext:  rename[col_ext]  = "Ingreso limpieza"
    if col_tot:  rename[col_tot]  = "Total ingresos"
    if col_port: rename[col_port] = "Portal"
    if col_comi: rename[col_comi] = "Comisión portal"

    out.rename(columns=rename, inplace=True)

    if "Ingreso limpieza" not in out.columns:
        out["Ingreso limpieza"] = 0.0

    for c in ["Ingreso alojamiento","Ingreso limpieza","Total ingresos","Comisión portal"]:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)

    for c in ["Fecha entrada","Fecha salida"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    if "Alojamiento" in out.columns:
        out["Alojamiento"] = out["Alojamiento"].astype(str).str.strip().str.upper()
        
        
        # -------------------------
    # HUESPEDES TOTALES
    # -------------------------

    col_adultos = first_existing(["Adultos","adultos"])
    col_ninos   = first_existing(["Niños","Ninos","niños","ninos"])

    if col_adultos:
        out["Adultos"] = pd.to_numeric(out[col_adultos], errors="coerce").fillna(0)
    else:
        out["Adultos"] = 0

    if col_ninos:
        out["Niños"] = pd.to_numeric(out[col_ninos], errors="coerce").fillna(0)
    else:
        out["Niños"] = 0

    out["Huéspedes totales"] = out["Adultos"] + out["Niños"]

    return out


# =====================================================
# MOTOR DE LIQUIDACIÓN DINÁMICO
# =====================================================

def process_dynamic(df, reglas):

    df = df.merge(reglas, left_on="Alojamiento", right_on="Property", how="left")
    df = df.loc[:, ~df.columns.duplicated()]

    # 🔥 SOLO apartamentos con reglas
    df = df[df["Property"].notna()].copy()

    if df.empty:
        st.warning("No hay apartamentos del CSV en el período seleccionado.")
        st.stop()

    # -------------------------
    # COMISIÓN PORTAL
    # -------------------------

    if "Portal" in df.columns:
        portal_series = df["Portal"]
        if isinstance(portal_series, pd.DataFrame):
            portal_series = portal_series.iloc[:, 0]
        mask_booking = portal_series.astype(str).str.lower().str.contains("booking", na=False)
    else:
        mask_booking = pd.Series(False, index=df.index)

    df["Comisión portal"] = pd.to_numeric(df["Comisión portal"], errors="coerce").fillna(0)

    df.loc[
        (mask_booking) & (df["skip_booking_vat"] == False),
        "Comisión portal"
    ] *= (1 + df["commission_vat_pct"] / 100)

    # -------------------------
    # IVA ALQUILER
    # -------------------------

    df["IVA del alquiler"] = 0.0

    mask_iva = df["compute_iva_alquiler"] == True

    df.loc[mask_iva, "IVA del alquiler"] = (
        df.loc[mask_iva, "Ingreso alojamiento"]
        - (df.loc[mask_iva, "Ingreso alojamiento"] / 1.10)
    )

    # -------------------------
    # HONORARIOS
    # -------------------------

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

    # -------------------------
    # LIMPIEZA
    # -------------------------

    df["Gasto limpieza"] = np.where(
        df["cleaning_fee"].notna(),
        df["cleaning_fee"],
        df["Ingreso limpieza"]
    )

    # -------------------------
    # AMENITIES
    # -------------------------

    df["Amenities"] = df["amenities_amount"].fillna(0)

    # -------------------------
    # TOTALES
    # -------------------------

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
        "Huéspedes totales",
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
    
    df = df.sort_values(by=["Alojamiento", "Fecha entrada"])

    return df[columnas_existentes]


# =====================================================
# EXPORTAR EXCEL
# =====================================================

def build_excel(df):

    wb = Workbook()
    ws = wb.active
    ws.title = "Liquidación"

    row_cursor = 1

    for aloj, subdf in df.groupby("Alojamiento"):

        # Título apartamento
        ws.cell(row=row_cursor, column=1, value=aloj).font = Font(bold=True)
        row_cursor += 1

        cols = list(subdf.columns)

        # Cabecera
        for col_idx, col in enumerate(cols, 1):
            ws.cell(row=row_cursor, column=col_idx, value=col).font = Font(bold=True)

        row_cursor += 1

        # Reservas
        for _, row in subdf.iterrows():
            for col_idx, col in enumerate(cols, 1):
                ws.cell(row=row_cursor, column=col_idx, value=row[col])
            row_cursor += 1

        # Fila TOTAL
        for col_idx, col in enumerate(cols, 1):
            if pd.api.types.is_numeric_dtype(subdf[col]):
                total_value = subdf[col].sum()
                ws.cell(row=row_cursor, column=col_idx, value=total_value).font = Font(bold=True)
            else:
                if col == "Fecha entrada":
                    ws.cell(row=row_cursor, column=col_idx, value="TOTAL").font = Font(bold=True)

        row_cursor += 2  # Espacio entre apartamentos

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
    
    # ---------------------------------
    # FILTRO POR ALOJAMIENTO
    # ---------------------------------

    alojamientos_disponibles = sorted(df_final["Alojamiento"].unique())

    alojamientos_seleccionados = st.multiselect(
    "Filtrar por alojamiento",
    options=alojamientos_disponibles,
    default=alojamientos_disponibles
    )

    df_final = df_final[df_final["Alojamiento"].isin(alojamientos_seleccionados)]


    st.success("Liquidación generada correctamente.")
    for aloj, subdf in df_final.groupby("Alojamiento"):
        st.subheader(f"🏠 {aloj}")

        block = subdf.copy()

        # Crear fila TOTAL
        total_row = {}

        for col in block.columns:
            if pd.api.types.is_numeric_dtype(block[col]):
                total_row[col] = block[col].sum()
            else:
                total_row[col] = ""

        total_row["Fecha entrada"] = "TOTAL"

        block = pd.concat([block, pd.DataFrame([total_row])], ignore_index=True)

        st.dataframe(block, use_container_width=True)
        st.divider()


    excel_file = build_excel(df_final)

    st.download_button(
        "Descargar Excel",
        excel_file.getvalue(),
        file_name="Liquidacion_dinamica.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
