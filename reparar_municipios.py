"""
Repara los nombres de municipios/departamentos dañados por codificación (Estelí -> 'EstelÃ\xad' -> 'ESTELA').

1) Reescribe la hoja Maestro_Municipio desde data/Maestro_Municipio.csv (UTF-8 correcto).
2) Corrige Municipio y Departamento en Registro_Precios (ESTELA -> ESTELI, SABACO -> SEBACO, ...).
3) Informa cuántas filas totalmente vacías hay en Registro_Precios.
   Con --borrar-vacias las elimina.

Uso (con valesK activo, dentro de precios_app):
    python reparar_municipios.py
    python reparar_municipios.py --borrar-vacias
"""
import sys
import tomllib
import unicodedata
from pathlib import Path

import gspread
import pandas as pd


def normalizar(txt: str) -> str:
    """Igual que en app.py: mayúsculas, sin tildes, sin espacios dobles."""
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return " ".join(t.upper().split())


def mojibake(txt: str) -> str:
    """Cómo queda un texto UTF-8 leído como Latin-1 (Estelí -> EstelÃ\xad)."""
    try:
        return txt.encode("utf-8").decode("latin-1")
    except UnicodeError:
        return txt


cfg = tomllib.load(open(".streamlit/secrets.toml", "rb"))
sh = gspread.service_account_from_dict(dict(cfg["gcp_service_account"])).open_by_key(cfg["SHEET_ID"])

# ── 1) Maestro_Municipio desde el CSV correcto ──
csv = Path("data/Maestro_Municipio.csv")
m = pd.read_csv(csv, dtype=str, encoding="utf-8").fillna("")
ws_m = sh.worksheet("Maestro_Municipio")
ws_m.clear()
ws_m.update(range_name="A1", values=[m.columns.tolist()] + m.values.tolist(), value_input_option="USER_ENTERED")
print(f"Maestro_Municipio reescrito: {len(m)} filas · deptos: {sorted(m['Departamento'].unique())}")

# ── 2) Diccionario dañado -> correcto (en forma normalizada, como guarda la app) ──
nombres = set(m["Municipio"]) | set(m["Departamento"])
mapa = {}
for n in nombres:
    malo, bueno = normalizar(mojibake(n)), normalizar(n)
    if malo != bueno:
        mapa[malo] = bueno

ws_r = sh.worksheet("Registro_Precios")
v = ws_r.get_all_values()
enc, filas = v[0], v[1:]
i_mun, i_dep = enc.index("Municipio"), enc.index("Departamento")
cambios = 0
for f in filas:
    for i in (i_mun, i_dep):
        if f[i] in mapa:
            f[i] = mapa[f[i]]
            cambios += 1
if cambios:
    col = lambda i: gspread.utils.rowcol_to_a1(1, i + 1)[:-1]
    ws_r.update(range_name=f"{col(i_mun)}2:{col(i_dep)}{len(filas) + 1}",
                values=[[f[i_mun], f[i_dep]] for f in filas])
print(f"Registro_Precios: {cambios} celdas corregidas")
print("  Municipios ahora:", sorted({f[i_mun] for f in filas if f[i_mun]}))
print("  Departamentos ahora:", sorted({f[i_dep] for f in filas if f[i_dep]}))

# ── 3) Filas vacías ──
vacias = [k + 2 for k, f in enumerate(filas) if not any(c.strip() for c in f)]
print(f"Filas totalmente vacías: {len(vacias)}" + (f" (primera: fila {vacias[0]})" if vacias else ""))
if vacias and "--borrar-vacias" in sys.argv:
    # agrupar en bloques contiguos y borrar de abajo hacia arriba (para no mover las demás)
    bloques, ini = [], vacias[0]
    for a, b in zip(vacias, vacias[1:] + [None]):
        if b != a + 1:
            bloques.append((ini, a))
            ini = b
    for ini, fin in reversed(bloques):
        ws_r.delete_rows(ini, fin)
    print(f"  Eliminadas {len(vacias)} filas vacías en {len(bloques)} bloque(s).")
print(f"Filas con datos: {sum(1 for f in filas if any(c.strip() for c in f))}")
