import pandas as pd
from pathlib import Path

# Ajusta estas rutas si hace falta
repo = Path(r"c:\Users\Usuario\Desktop\liquidaciones-pp")
rules_fp = repo / "reglas_apartamentos.csv"
reservas_fp = repo / "Reservas enero - copia.csv"

# Carga reglas (CSV coma-decimal punto)
rules = pd.read_csv(rules_fp, dtype=str).fillna("")
rules_props = set(rules["property"].astype(str).str.strip().str.upper())

# Carga reservas (archivo CSV con separador ';' y decimales ',')
encodings = ['utf-8', 'cp1252', 'latin-1']
res = None
last_exc = None
for enc in encodings:
    try:
        res = pd.read_csv(reservas_fp, sep=';', decimal=',', encoding=enc, dtype=str).fillna("")
        print(f"Leído reservas con encoding: {enc}")
        break
    except Exception as e:
        last_exc = e
# fallback seguro
if res is None:
    res = pd.read_csv(reservas_fp, sep=';', decimal=',', encoding='latin-1', dtype=str, errors='replace').fillna("")
    print("Fallback: leído con latin-1 y errors='replace' (caracteres inválidos reemplazados)")

# Normalizar columna de alojamiento (puede tener nombre distinto en tu CSV)
# Busca columna que contenga "Nombre alojamiento" o "Nombre alojamiento" parecido
possible_cols = [c for c in res.columns if "aloj" in c.lower()]
if not possible_cols:
    print("No pude encontrar la columna de alojamiento en las reservas. Columnas disponibles:", list(res.columns)[:20])
else:
    aloj_col = possible_cols[0]
    reservas_props = set(res[aloj_col].astype(str).str.strip().str.upper())
    missing = sorted([p for p in reservas_props if p and p not in rules_props])
    print(f"Propiedades en 'reglas_apartamentos.csv': {len(rules_props)}")
    print(f"Propiedades detectadas en reservas ({aloj_col}): {len(reservas_props)}")
    print("\nPropiedades en reservas que NO están en reglas_apartamentos.csv (ejemplos):")
    for p in missing[:80]:
        print("-", p)
    if not missing:
        print("✅ Todas las propiedades de las reservas están cubiertas por el CSV de reglas.")