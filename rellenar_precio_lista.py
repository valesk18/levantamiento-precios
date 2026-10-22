"""
Completa Precio_Entrada_Lista en Registro_Precios para los SKU propios (FANNY) que quedaron vacíos,
usando el precio de lista de Maestro_SKU según la zona de la ruta. No toca ninguna otra columna.

Uso (con valesK activo, dentro de precios_app):
    python rellenar_precio_lista.py
"""
import tomllib

import gspread

ZONAS = [("MA_", "MANAGUA"), ("HORECA", "MANAGUA"), ("MERCADOS", "MANAGUA"), ("OFICINA", "MANAGUA"),
         ("SUPER", "MANAGUA"), ("CENTRO", "CENTRO"), ("NORTE", "NORTE"),
         ("OCCIDENTE", "OCCIDENTE_SURORIENTE"), ("SURORIENTE", "OCCIDENTE_SURORIENTE")]
CLAVES = {"MANAGUA": ["MANAGUA"], "CENTRO": ["CENTRO"], "NORTE": ["NORTE"],
          "OCCIDENTE_SURORIENTE": ["CIDENTE", "SURORIENTE", "SUR_ORIENTE"]}

cfg = tomllib.load(open(".streamlit/secrets.toml", "rb"))
sh = gspread.service_account_from_dict(dict(cfg["gcp_service_account"])).open_by_key(cfg["SHEET_ID"])

sku = sh.worksheet("Maestro_SKU").get_all_records()
cols = list(sku[0].keys()) if sku else []


def columna(zona):
    for c in cols:
        cu = str(c).upper().replace(" ", "_")
        if cu.startswith("PRECIO_") and cu != "PRECIO_SUGERIDO" and any(k in cu for k in CLAVES.get(zona, [zona])):
            return c
    return ""


col_zona = {z: columna(z) for z in CLAVES}
print("Columnas de precio por zona:", col_zona)

ws = sh.worksheet("Registro_Precios")
v = ws.get_all_values()
enc, filas = v[0], v[1:]
iR, iS, iT, iL = (enc.index(c) for c in ("Ruta", "SKU_ID", "Tipo", "Precio_Entrada_Lista"))

lista = []
llenadas = 0
for f in filas:
    actual = f[iL] if len(f) > iL else ""
    nuevo = actual
    if f[iT] == "Propio" and not str(actual).strip():
        zona = next((z for p, z in ZONAS if f[iR].upper().startswith(p)), "")
        c = col_zona.get(zona, "")
        p = next((str(s.get(c, "")).strip() for s in sku if s["SKU_ID"] == f[iS]), "") if c else ""
        if p:
            nuevo, llenadas = p, llenadas + 1
    lista.append([nuevo])

col = gspread.utils.rowcol_to_a1(1, iL + 1)[:-1]
ws.update(range_name=f"{col}2:{col}{len(filas) + 1}", values=lista, value_input_option="USER_ENTERED")
print(f"Precio_Entrada_Lista completado en {llenadas} filas FANNY.")
