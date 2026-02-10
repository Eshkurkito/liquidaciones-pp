import pandas as pd
import numpy as np
from utils import normalize_columns, ensure_required, NIGHTS_COL

def _cell_value(row: pd.Series, key, default=""):
    val = row.get(key, default)
    if isinstance(val, (pd.Series, np.ndarray, list, tuple)):
        return default
    if pd.isna(val):
        return default
    return val

def process_by_rules(df: pd.DataFrame, rules_map: dict, default_commission_vat: float = 21.0):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Total ingresos","Comisión portal","Portal"], "Procesar por reglas")

    out = df.copy()
    out["Ingreso alojamiento"] = pd.to_numeric(out.get("Ingreso alojamiento",0.0), errors="coerce").fillna(0.0)
    out["Ingreso limpieza"] = pd.to_numeric(out.get("Ingreso limpieza",0.0), errors="coerce").fillna(0.0)
    out["Comisión portal"] = pd.to_numeric(out.get("Comisión portal",0.0), errors="coerce").fillna(0.0)
    out["Total ingresos"] = pd.to_numeric(out.get("Total ingresos",0.0), errors="coerce").fillna(0.0)

    def compute_row(r):
        prop = str(_cell_value(r, "Alojamiento", "")).strip().upper()
        rule = rules_map.get(prop, {})

        ingreso = float(_cell_value(r, "Ingreso alojamiento", 0.0))
        com_orig = float(_cell_value(r, "Comisión portal", 0.0))
        portal = str(_cell_value(r, "Portal", "")).strip().lower()

        honorarios_pct = float(rule.get("honorarios_pct") or 0.20)
        honorarios_apply_vat = bool(int(rule.get("honorarios_apply_vat") or 1))
        honorarios_vat_pct = float(rule.get("honorarios_vat_pct") or 21.0)
        amenities_amount = float(rule.get("amenities_amount") or 0.0)

        cleaning_fee_val = rule.get("cleaning_fee")
        if cleaning_fee_val is None:
            cleaning_fee = float(_cell_value(r, "Ingreso limpieza", 0.0))
        else:
            try:
                cleaning_fee = float(cleaning_fee_val)
            except Exception:
                cleaning_fee = float(_cell_value(r, "Ingreso limpieza", 0.0))

        compute_iva_alquiler = bool(int(rule.get("compute_iva_alquiler") or 0))
        commission_vat_pct = float(rule.get("commission_vat_pct") if rule.get("commission_vat_pct") not in (None,"") else default_commission_vat)
        treat_empty = bool(int(rule.get("treat_empty_portal_as_booking") or 0))
        skip_booking = bool(int(rule.get("skip_booking_vat") or 0))
        split_comm = bool(int(rule.get("split_commission") or 0))
        hon_base_excl_com = bool(int(rule.get("hon_base_exclude_commission") or 0))

        is_booking = "booking" in portal
        is_empty = portal == ""

        com_sin_iva = com_orig
        iva_com = 0.0
        com_total = com_orig

        if not split_comm:
            if (is_booking or (is_empty and treat_empty)) and commission_vat_pct > 0 and (not skip_booking):
                iva_com = com_orig * (commission_vat_pct / 100.0)
                com_total = com_orig + iva_com
            else:
                com_total = com_orig
        else:
            iva_com = com_orig * (commission_vat_pct / 100.0) if commission_vat_pct > 0 else 0.0
            com_total = com_sin_iva + iva_com

        iva_alq = (ingreso - (ingreso / 1.10)) if compute_iva_alquiler else 0.0

        base_hon = ingreso
        if hon_base_excl_com:
            base_hon = ingreso - com_sin_iva

        if honorarios_apply_vat and honorarios_vat_pct:
            honorarios = base_hon * honorarios_pct * (1 + honorarios_vat_pct/100.0)
        else:
            honorarios = base_hon * honorarios_pct

        gasto_limpieza = cleaning_fee
        total_gastos = round(com_total + honorarios + gasto_limpieza + amenities_amount, 2)
        pago_prop = round(float(_cell_value(r, "Total ingresos", 0.0)) - total_gastos, 2)
        pago_recibido = round(float(_cell_value(r, "Total ingresos", 0.0)) - com_total, 2)

        res = {
            "Comisión portal": round(com_total, 2),
            "Honorarios Florit": round(honorarios, 2),
            "Gasto limpieza": round(gasto_limpieza, 2),
            "Amenities": round(amenities_amount, 2),
            "Total Gastos": total_gastos,
            "Pago al propietario": pago_prop,
            "Pago recibido": pago_recibido
        }

        res["IVA del alquiler"] = round(iva_alq, 2) if compute_iva_alquiler else None

        if split_comm:
            res["Comisión portal (sin IVA)"] = round(com_sin_iva, 2)
            res["IVA comisión portal"] = round(iva_com, 2)

        return pd.Series(res)

    computed = out.apply(compute_row, axis=1)
    out.update(computed)

    expected_calc_cols = [
        "Comisión portal (sin IVA)", "IVA comisión portal", "Comisión portal",
        "Honorarios Florit", "Gasto limpieza", "Amenities",
        "Total Gastos", "Pago al propietario", "Pago recibido", "IVA del alquiler"
    ]
    for col in expected_calc_cols:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0).round(2)

    for c in out.columns:
        if c != NIGHTS_COL and pd.api.types.is_numeric_dtype(out[c]):
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).round(2)

    # Ensure Portal and Comisión portal are treated as Series (handle duplicate column names)
    if "Portal" in out.columns:
        portal_series = out["Portal"]
    else:
        portal_series = pd.Series([""] * len(out), index=out.index)

    com_series = out["Comisión portal"] if "Comisión portal" in out.columns else pd.Series([0.0] * len(out), index=out.index)

    portal_stripped = portal_series.astype(str).str.strip()
    com_numeric = pd.to_numeric(com_series, errors="coerce").fillna(0.0)

    warn = int(((portal_stripped == "") & (com_numeric > 0)).sum())

    return out, warn