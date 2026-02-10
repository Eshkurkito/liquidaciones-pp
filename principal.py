import streamlit as st
import pandas as pd
import numpy as np
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import re
from pathlib import Path
from typing import Optional, Dict

st.set_page_config(page_title="LIQUIDACIONES (CSV rules)", page_icon="🏦", layout="wide")

# ---------- Constantes / utilidades ----------
MONEY_COLS_CANON = {
    "Ingreso alojamiento","Ingreso limpieza","Total ingresos","Comisión portal",
    "Honorarios Florit","Gasto limpieza","Amenities","Total Gastos","Pago al propietario",
    "Pago recibido","IVA del alquiler"
}
NIGHTS_COL = "Noches ocupadas"

def ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    counts, new_cols = {}, []
    for c in df.columns:
        name = str(c)
        n = counts.get(name, 0)
        new_cols.append(name if n == 0 else f"{name}.{n}")
        counts[name] = n + 1
    out = df.copy()
    out.columns = new_cols
    return out

def _first_existing(df, candidates):
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = str(cand).strip().lower()
        if k in norm_map:
            return norm_map[k]
    return None

def ensure_required(df, required, ctx=""):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas requeridas: {missing} en {ctx}. Ajusta el archivo o usa el modo por letras.")
        st.stop()

def base_name(colname: str) -> str:
    return re.sub(r"\.\d+$", "", str(colname)).strip()

def is_money_col(colname: str) -> bool:
    return base_name(colname) in MONEY_COLS_CANON

def is_nights_col(colname: str) -> bool:
    return base_name(colname).lower() == NIGHTS_COL.lower()

def fmt_number_for_ui(colname: str, x):
    if is_nights_col(colname):
        try: return f"{int(round(float(x)))}"
        except Exception: return x
    if is_money_col(colname):
        try:
            s = f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"{s} €"
        except Exception:
            return x
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

def find_col(df: pd.DataFrame, base: str):
    for c in df.columns:
        if base_name(c).lower() == base.strip().lower():
            return c
    return None

def show_table_es_grouped(df: pd.DataFrame, title: str, group_col: str = "Alojamiento"):
    st.subheader(title)
    if group_col not in df.columns:
        view = df.copy()
        total = {}
        for c in view.columns:
            if pd.api.types.is_numeric_dtype(view[c]):
                total[c] = view[c].sum()
            elif pd.api.types.is_datetime64_any_dtype(view[c]):
                total[c] = pd.NaT
            else:
                total[c] = ""
        view = pd.concat([view, pd.DataFrame([total], index=["TOTAL"])], axis=0)
        view_fmt = view.copy()
        for c in view_fmt.columns:
            if pd.api.types.is_numeric_dtype(view[c]) or is_money_col(c) or is_nights_col(c):
                view_fmt[c] = view_fmt[c].apply(lambda v: fmt_number_for_ui(c, v))
        def highlight_total(row):
            return ["font-weight: bold;" if row.name == "TOTAL" else "" for _ in row]
        st.dataframe(view_fmt.style.apply(highlight_total, axis=1), use_container_width=True)
        return
    for aloj, subdf in df.groupby(group_col):
        st.markdown(f"**{aloj}**")
        block = subdf.copy()
        total = {}
        for c in block.columns:
            if pd.api.types.is_numeric_dtype(block[c]):
                total[c] = block[c].sum()
            elif pd.api.types.is_datetime64_any_dtype(block[c]):
                total[c] = pd.NaT
            else:
                total[c] = ""
        block = pd.concat([block, pd.DataFrame([total], index=["TOTAL"])], axis=0)
        block_fmt = block.copy()
        for c in block_fmt.columns:
            if pd.api.types.is_numeric_dtype(block[c]) or is_money_col(c) or is_nights_col(c):
                block_fmt[c] = block_fmt[c].apply(lambda v: fmt_number_for_ui(c, v))
        def highlight_total(row):
            return ["font-weight: bold;" if row.name == "TOTAL" else "" for _ in row]
        st.dataframe(block_fmt.style.apply(highlight_total, axis=1), use_container_width=True)
        st.divider()

# ---------- Normalización de columnas ----------
LETTER_MAP_DEFAULT = {
    "W": "Alojamiento",
    "D": "Fecha entrada",
    "F": "Fecha salida",
    "H": "Noches ocupadas",
    "I": "Ingreso alojamiento",
    "L": "Ingreso limpieza",
    "O": "Total ingresos",
    "AP": "Portal",
    "AR": "Comisión portal",
    "AL": "IVA del alquiler",
}

def letters_to_idx(letter):
    s = letter.upper(); n = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'): return None
        n = n*26 + (ord(ch)-ord('A')+1)
    return n-1

def normalize_columns_by_letters(df, letter_map=LETTER_MAP_DEFAULT):
    out = df.copy(); cols = list(out.columns); rename = {}
    for L, std in letter_map.items():
        i = letters_to_idx(L)
        if i is not None and i < len(cols):
            rename[cols[i]] = std
    out.rename(columns=rename, inplace=True)
    return normalize_columns(out)

def normalize_columns(df):
    out = df.copy()

    def clean_money_series(ser: pd.Series) -> pd.Series:
        s = ser.astype(str).fillna("")
        # quitar todo menos dígitos, coma, punto y signo -
        s = s.str.replace(r"[^\d,.\-]", "", regex=True)
        has_dot = s.str.contains(r"\.", regex=True)
        has_comma = s.str.contains(r",", regex=True)
        # ambos presentes -> asumimos formato ES: '.' miles, ',' decimal -> quitar '.' y cambiar ','->'.'
        both = has_dot & has_comma
        s.loc[both] = s.loc[both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        # solo coma -> coma decimal
        only_comma = (~has_dot) & has_comma
        s.loc[only_comma] = s.loc[only_comma].str.replace(",", ".", regex=False)
        # vacíos a 0
        s = s.replace("", "0")
        return pd.to_numeric(s, errors="coerce").fillna(0.0)

    col_aloj = _first_existing(out, ["Nombre alojamiento","Alojamiento","Nombre del alojamiento","Nombre Alojamiento"])
    col_fent = _first_existing(out, ["Fecha entrada","Fecha de entrada"])
    col_fsal = _first_existing(out, ["Fecha salida","Fecha de salida"])
    col_noch = _first_existing(out, ["Noches","noches","Noches ocupadas"])
    col_alq  = _first_existing(out, ["Alquiler con tasas","Ingreso alojamiento","Importe alojamiento"])
    col_ext  = _first_existing(out, ["Ingreso limpieza","Tarifa limpieza","Limpieza","Importe limpieza","Extras con tasas","Gastos de limpieza","Gasto limpieza"])
    col_tot  = _first_existing(out, ["Total reserva con tasas","Total ingresos","Total"])
    col_port = _first_existing(out, ["Web origen","Portal","Canal","Fuente"])
    col_comi = _first_existing(out, ["Comisión Portal/Intermediario: Comisión calculada","Comisión portal","Comisión"])
    col_ivaal= _first_existing(out, ["IVA del alojamiento","IVA alojamiento","IVA del alquiler"])

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
    if col_ivaal:rename[col_ivaal]= "IVA del alquiler"

    out.rename(columns=rename, inplace=True)

    # Limpiar y tipar correctamente columnas monetarias (gestiona "1.234,56 €", "1234,56", "1234.56", etc.)
    for c in ["Ingreso alojamiento","Ingreso limpieza","Total ingresos","Comisión portal","IVA del alquiler"]:
        if c in out.columns:
            out[c] = clean_money_series(out[c])

    # Noches y fechas
    if "Noches ocupadas" in out.columns:
        out["Noches ocupadas"] = pd.to_numeric(out["Noches ocupadas"].astype(str).str.replace(r"[^\d\-]", "", regex=True), errors="coerce").fillna(0).round(0).astype(int)
    for c in ["Fecha entrada","Fecha salida"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    if "Alojamiento" in out.columns:
        out["Alojamiento"] = out["Alojamiento"].astype(str).str.strip().str.upper()

    return out

# ---------- Carga de reglas desde CSV ----------
def load_rules_csv(path: Optional[str] = None) -> Dict[str, dict]:
    p = Path(path) if path else Path(__file__).resolve().parent / "reglas_apartamentos.csv"
    if not p.exists():
        return {}
    # leer como texto y normalizar
    df = pd.read_csv(p, dtype=str).fillna("")
    df["property_norm"] = df["property"].astype(str).str.strip().str.upper()
    def conv(x):
        try:
            if x == "": return None
            return float(x)
        except Exception:
            return x
    numeric_cols = ["honorarios_pct","honorarios_vat_pct","amenities_amount","cleaning_fee","commission_vat_pct"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(conv)
    int_cols = ["honorarios_apply_vat","compute_iva_alquiler","treat_empty_portal_as_booking",
                "skip_booking_vat","split_commission","hon_base_exclude_commission"]
    for col in int_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).replace({"1":"1","0":"0","": "0"}).astype(int)
    rules = df.set_index("property_norm").to_dict(orient="index")
    return rules

RULES_MAP = load_rules_csv(None)

# ---------- Procesador central: usa reglas por piso (CSV) ----------
def process_by_rules(df: pd.DataFrame, rules_map: dict, default_commission_vat: float = 21.0):
    """
    Procesa fila a fila usando las reglas por apartment (rules_map).
    Reproduce las fórmulas del BACK UP pero leyendo parámetros desde reglas_apartamentos.csv.
    """
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Total ingresos","Comisión portal","Portal"], "Procesar por reglas")

    out = df.copy()
    # asegurar tipos básicos
    out["Ingreso alojamiento"] = pd.to_numeric(out.get("Ingreso alojamiento",0.0), errors="coerce").fillna(0.0)
    out["Ingreso limpieza"] = pd.to_numeric(out.get("Ingreso limpieza",0.0), errors="coerce").fillna(0.0)
    out["Comisión portal"] = pd.to_numeric(out.get("Comisión portal",0.0), errors="coerce").fillna(0.0)
    out["Total ingresos"] = pd.to_numeric(out.get("Total ingresos",0.0), errors="coerce").fillna(0.0)

    def compute_row(r):
        prop = str(r.get("Alojamiento","")).strip().upper()
        rule = rules_map.get(prop, {})

        ingreso = float(r.get("Ingreso alojamiento", 0.0))
        com_orig = float(r.get("Comisión portal", 0.0))
        portal = str(r.get("Portal","") or "").strip().lower()

        # leer parámetros desde la regla (con defaults)
        honorarios_pct = float(rule.get("honorarios_pct") or 0.20)
        honorarios_apply_vat = bool(int(rule.get("honorarios_apply_vat") or 1))
        honorarios_vat_pct = float(rule.get("honorarios_vat_pct") or 21.0)
        amenities_amount = float(rule.get("amenities_amount") or 0.0)

        # cleaning_fee: si la regla especifica vacío o 0 -> usar "Ingreso limpieza" de la fila
        cleaning_fee_raw = rule.get("cleaning_fee")
        try:
            cleaning_fee_val = float(cleaning_fee_raw) if cleaning_fee_raw not in (None, "") else None
        except Exception:
            cleaning_fee_val = None
        if cleaning_fee_val in (None, 0.0):
            cleaning_fee = float(r.get("Ingreso limpieza", 0.0))
        else:
            cleaning_fee = cleaning_fee_val

        compute_iva_alquiler = bool(int(rule.get("compute_iva_alquiler") or 0))
        commission_vat_pct = float(rule.get("commission_vat_pct") if rule.get("commission_vat_pct") not in (None,"") else default_commission_vat)
        treat_empty = bool(int(rule.get("treat_empty_portal_as_booking") or 0))
        skip_booking = bool(int(rule.get("skip_booking_vat") or 0))
        split_comm = bool(int(rule.get("split_commission") or 0))
        hon_base_excl_com = bool(int(rule.get("hon_base_exclude_commission") or 0))

        # ---- cálculo comisión (mismas reglas que BACK UP) ----
        is_booking = "booking" in portal
        is_empty = portal == ""