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
                pd.to_numeric(reglas[col], errors="coerce")
                .fillna(0)
                .astype(int)
                .map({1: True, 0: False, 2: False})
            )

            
    if "self_managed" in reglas.columns:
        reglas["self_managed"] = (
        pd.to_numeric(reglas["self_managed"], errors="coerce")
        .fillna(0)
        .astype(int)
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
    
    # Buscar columna comisión de forma flexible
    # Detectar columna de comisión de forma flexible
    col_comi = None

    for col in out.columns:
        nombre = str(col).lower()
        if "comisión" in nombre or "comision" in nombre:
            col_comi = col
            break

    if col_comi:
        out.rename(columns={col_comi: "Comisión portal"}, inplace=True)



    rename = {}
    if col_aloj: rename[col_aloj] = "Alojamiento"
    if col_fent: rename[col_fent] = "Fecha entrada"
    if col_fsal: rename[col_fsal] = "Fecha salida"
    if col_noch: rename[col_noch] = "Noches ocupadas"
    if col_alq:  rename[col_alq]  = "Ingreso alojamiento"
    if col_ext:  rename[col_ext]  = "Ingreso limpieza"
    if col_tot:  rename[col_tot]  = "Total ingresos"
    if col_port: rename[col_port] = "Portal"
    if col_comi: 
        rename[col_comi] = "Comisión portal"

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

    # Garantizar que exista Comisión portal
    if "Comisión portal" not in out.columns:
        out["Comisión portal"] = 0.0

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
    # HONORARIOS (VERSIÓN RÁPIDA)
    # -------------------------

    base = df["Ingreso alojamiento"].copy()

    # Excluir comisión si aplica
    mask_exclude = df["hon_base_exclude_commission"] == True
    base = np.where(mask_exclude, base - df["Comisión portal"], base)

    # Excluir IVA si aplica
    mask_iva_base = df["compute_iva_alquiler"] == True
    base = np.where(mask_iva_base, base - df["IVA del alquiler"], base)

    # Calcular honorarios base
    df["Honorarios Florit"] = base * df["honorarios_pct"]

    # Aplicar IVA a honorarios si corresponde
    mask_vat = df["honorarios_apply_vat"] == True
    df.loc[mask_vat, "Honorarios Florit"] *= (
        1 + df.loc[mask_vat, "honorarios_vat_pct"] / 100
    )


    # -------------------------
    # LIMPIEZA
    # -------------------------

    df["Gasto limpieza"] = np.where(
        df["cleaning_fee"] > 0,
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
    
    # ---------------------------------
    # SELF MANAGED (Florit = propietario)
    # ---------------------------------

    # ---------------------------------
    # SELF MANAGED (Florit = propietario)
    # ---------------------------------

    
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
    
    if "self_managed" in df.columns:

        mask_self = df["self_managed"] == 1

        # Asegurar que todo es numérico
        cols_needed = [
            "Total ingresos",
            "Comisión portal",
            "Gasto limpieza",
            "Amenities"
        ]

        for col in cols_needed:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Honorarios = beneficio real
        df.loc[mask_self, "Honorarios Florit"] = (
            df.loc[mask_self, "Total ingresos"]
            - df.loc[mask_self, "Comisión portal"]
            - df.loc[mask_self, "Gasto limpieza"]
            - df.loc[mask_self, "Amenities"]
        )

        # Total Gastos = solo gastos reales (sin honorarios)
        df.loc[mask_self, "Total Gastos"] = (
            df.loc[mask_self, "Comisión portal"]
            + df.loc[mask_self, "Gasto limpieza"]
            + df.loc[mask_self, "Amenities"]
        )

        # Pago propietario = mismo importe que honorarios
        df.loc[mask_self, "Pago al propietario"] = df.loc[mask_self, "Honorarios Florit"]

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



st.title("📊 Liquidaciones dinámicas")

# 🔹 TABS DESPUÉS


# INPUTS
start_date = st.date_input("Desde", key="desde_global")
end_date = st.date_input("Hasta", key="hasta_global")
file_reservas = st.file_uploader("Sube archivo (.xlsx)", type=["xlsx"], key="file_global")
df_base = None

if file_reservas:
    df_base = pd.read_excel(file_reservas)
    df_base = normalize_columns(df_base)


st.divider()

tab1, tab2 = st.tabs(["Liquidaciones", "Previsión Tesorería"])

reglas = load_reglas()

if "df_final" not in st.session_state:
    st.session_state.df_final = None

st.divider()  # Opcional para separar visualmente





with tab1:
    
    

    # ---------------------------------
    # CARGAR REGLAS Y CREAR FILTRO
    # ---------------------------------


    alojamientos_gestionados = sorted(reglas["Property"].unique())

    # Inicializar estado si no existe
    if "alojamientos_sel" not in st.session_state:
        st.session_state.alojamientos_sel = alojamientos_gestionados

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Seleccionar todos"):
            st.session_state.alojamientos_sel = alojamientos_gestionados

    with col2:
        if st.button("Quitar todos"):
            st.session_state.alojamientos_sel = []

    alojamientos_seleccionados = st.multiselect(
    "Selecciona alojamiento(s)",
    options=alojamientos_gestionados,
    default=st.session_state.alojamientos_sel,
    key="alojamientos_sel"
    )


    # ---------------------------------
    # BOTÓN
    # ---------------------------------

    if st.button("Generar liquidación"):

        if not file_reservas:
            st.error("Sube el archivo de reservas.")
            st.stop()

        df_res = df_base.copy()


        # 🔥 FILTRAR ANTES DEL CÁLCULO
        df_res = df_res[df_res["Alojamiento"].isin(alojamientos_seleccionados)]

        df_res = df_res[
            (df_res["Fecha entrada"] >= pd.to_datetime(start_date))
            & (df_res["Fecha entrada"] <= pd.to_datetime(end_date))
        ]

        if df_res.empty:
            st.warning("No hay reservas para los alojamientos seleccionados.")
            st.stop()

        st.session_state.df_final = process_dynamic(df_res, reglas)

    
        # ---------------------------------
        # PISOS SIN LIQUIDACIÓN
        # ---------------------------------

        pisos_gestionados = set(reglas["Property"].unique())
        pisos_con_liquidacion = set(st.session_state.df_final["Alojamiento"].unique())

        pisos_sin_liquidacion = sorted(pisos_gestionados - pisos_con_liquidacion)

        if pisos_sin_liquidacion:
            st.info(
            f"Pisos sin liquidación en este período: {', '.join(pisos_sin_liquidacion)}"
        )
        else:
        
            st.success("Todos los pisos gestionados tienen liquidación en este período.")


        st.success("Liquidación generada correctamente.")

        for aloj, subdf in st.session_state.df_final.groupby("Alojamiento"):

            st.subheader(f"🏠 {aloj}")

            block = subdf.copy()

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

        excel_file = build_excel(st.session_state.df_final)

        st.download_button(
            "Descargar Excel",
            excel_file.getvalue(),
            file_name="Liquidacion_dinamica.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# =====================================================
# TAB 2 → PREVISIÓN TESORERÍA
# =====================================================

with tab2:

    st.header("📈 Previsión Tesorería – Honorarios Florit")

    # 🔹 Verificar archivo
    if df_base is None:
        st.info("Sube archivo para ver previsión.")
        st.stop()

    # 🔹 Filtro independiente alojamientos
    alojamientos_tab2 = st.multiselect(
        "Filtrar alojamiento(s)",
        options=reglas["Property"].unique(),
        key="alojamientos_tab2"
    )

    fecha_corte = st.date_input(
        "Fecha de corte",
        value=end_date,
        key="fecha_corte_tesoreria"
    )
    
    calcular_tesoreria = st.button("Calcular previsión")

    if calcular_tesoreria:

        df_periodo = df_base.copy()

        # Aplicar filtro alojamientos
        if alojamientos_tab2:
            df_periodo = df_periodo[
                df_periodo["Alojamiento"].isin(alojamientos_tab2)
            ]

        # Aplicar filtro fechas
        df_periodo = df_periodo[
            (df_periodo["Fecha entrada"] >= pd.to_datetime(start_date)) &
            (df_periodo["Fecha entrada"] <= pd.to_datetime(end_date))
        ]

        # Calcular honorarios
        df_periodo = process_dynamic(df_periodo, reglas)

        if df_periodo.empty:
            st.warning("No hay datos para el periodo seleccionado.")
            st.stop()

        # Guardamos en session_state para no recalcular
        st.session_state.df_tesoreria = df_periodo
        
    if "df_tesoreria" in st.session_state:

        df_periodo = st.session_state.df_tesoreria


        
    # =====================================================
    # COMPARACIÓN AÑO ACTUAL VS AÑO ANTERIOR (YTD)
    # =====================================================

    anio_actual = end_date.year
    anio_anterior = anio_actual - 1

    # Mismo rango de fechas pero año anterior
    fecha_inicio_anterior = start_date.replace(year=anio_anterior)
    fecha_fin_anterior = end_date.replace(year=anio_anterior)

    df_anterior = df_base.copy()

    # Aplicar filtro alojamientos
    if alojamientos_tab2:
        df_anterior = df_anterior[
            df_anterior["Alojamiento"].isin(alojamientos_tab2)
        ]

    # Aplicar rango equivalente año anterior
    df_anterior = df_anterior[
        (df_anterior["Fecha entrada"] >= pd.to_datetime(fecha_inicio_anterior)) &
        (df_anterior["Fecha entrada"] <= pd.to_datetime(fecha_fin_anterior))
    ]

    df_anterior = process_dynamic(df_anterior, reglas)

    hon_actual = df_periodo["Honorarios Florit"].sum()
    hon_anterior = df_anterior["Honorarios Florit"].sum()

    variacion = hon_actual - hon_anterior
    variacion_pct = (
    variacion / hon_anterior * 100
    if hon_anterior > 0 else 0
    )


    # 🔹 Métricas
    df_corte = df_periodo[
        df_periodo["Fecha entrada"] <= pd.to_datetime(fecha_corte)
    ]

    honorarios_corte = df_corte["Honorarios Florit"].sum()
    honorarios_periodo = df_periodo["Honorarios Florit"].sum()

    porcentaje = (
        honorarios_corte / honorarios_periodo * 100
        if honorarios_periodo > 0 else 0
    )

    pendiente = honorarios_periodo - honorarios_corte

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("💶 Generado hasta corte", f"{honorarios_corte:,.2f} €")
    col2.metric("📅 Total periodo", f"{honorarios_periodo:,.2f} €")
    col3.metric("📊 % ejecutado", f"{porcentaje:,.1f} %")
    col4.metric("🔮 Pendiente periodo", f"{pendiente:,.2f} €")
    
    st.divider()

    col5, col6, col7 = st.columns(3)

    col5.metric(f"💶 {anio_actual}", f"{hon_actual:,.2f} €")
    col6.metric(f"💶 {anio_anterior}", f"{hon_anterior:,.2f} €")
    col7.metric("📈 Variación %", f"{variacion_pct:,.1f} %")


    st.divider()

    ranking = (
        df_periodo
        .groupby("Alojamiento")["Honorarios Florit"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    st.subheader("🏠 Ranking por apartamento")
    st.dataframe(ranking, use_container_width=True)
