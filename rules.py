import os
import pandas as pd

def load_rules_csv(path: str = "reglas_apartamentos.csv") -> dict:
    """
    Lee reglas desde CSV y devuelve mapping PROPERTY -> rule dict.
    Campos esperados en el CSV: property,honorarios_pct,honorarios_apply_vat,
    honorarios_vat_pct,amenities_amount,cleaning_fee,compute_iva_alquiler,
    commission_vat_pct,treat_empty_portal_as_booking,skip_booking_vat,
    split_commission,hon_base_exclude_commission,notes
    """
    if not os.path.exists(path):
        return {}
    df = pd.read_csv(path, dtype=str).fillna("")
    rules = {}
    for _, row in df.iterrows():
        prop = str(row.get("property","")).strip().upper()
        if not prop:
            continue

        def fnum(k, default=0.0):
            v = row.get(k,"")
            try:
                return float(v) if v != "" else default
            except Exception:
                return default

        def fint(k, default=0):
            v = row.get(k,"")
            try:
                return int(float(v)) if v != "" else default
            except Exception:
                return default

        cleaning_raw = row.get("cleaning_fee","")
        cleaning_val = None
        try:
            if cleaning_raw != "":
                cleaning_val = float(cleaning_raw)
                # tratar 0.0 en CSV como "usar valor de la fila"
                if cleaning_val == 0.0:
                    cleaning_val = None
        except Exception:
            cleaning_val = None

        rules[prop] = {
            "honorarios_pct": fnum("honorarios_pct", 0.2),
            "honorarios_apply_vat": fint("honorarios_apply_vat", 1),
            "honorarios_vat_pct": fnum("honorarios_vat_pct", 21.0),
            "amenities_amount": fnum("amenities_amount", 0.0),
            "cleaning_fee": cleaning_val,
            "compute_iva_alquiler": fint("compute_iva_alquiler", 0),
            "commission_vat_pct": fnum("commission_vat_pct", 21.0),
            "treat_empty_portal_as_booking": fint("treat_empty_portal_as_booking", 0),
            "skip_booking_vat": fint("skip_booking_vat", 0),
            "split_commission": fint("split_commission", 0),
            "hon_base_exclude_commission": fint("hon_base_exclude_commission", 0),
            "notes": row.get("notes","")
        }
    return rules