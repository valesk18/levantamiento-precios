"""
Para los SKU propios (FANNY) ya guardados en Registro_Precios:
- completa Precio_Entrada_Lista con el precio de lista de Maestro_SKU según la zona de la ruta;
- si Precio_Entrada está vacío (y hay existencia), lo llena con ese precio de lista y recalcula el margen.
No toca Precio_Salida ni las filas de competencia.

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
ix = {c: enc.index(c) for c in ("Ruta", "SKU_ID", "Tipo", "Disponible", "Precio_Entrada", "Precio_Salida",
                                 "Margen_Abs", "Margen_Pct", "Precio_Entrada_Lista")}


def num(x):
    try:
        return float(str(x).replace(",", "").strip())
    except ValueError:
        return None


n_lista = n_entrada = 0
for f in filas:
    f += [""] * (len(enc) - len(f))
    if f[ix["Tipo"]] != "Propio":
        continue
    zona = next((z for p, z in ZONAS if f[ix["Ruta"]].upper().startswith(p)), "")
    c = col_zona.get(zona, "")
    p = next((str(s.get(c, "")).strip() for s in sku if s["SKU_ID"] == f[ix["SKU_ID"]]), "") if c else ""
    if p and not f[ix["Precio_Entrada_Lista"]].strip():
        f[ix["Precio_Entrada_Lista"]] = p
        n_lista += 1
    # Entrada vacía y producto con existencia -> se usa la de lista
    if f[ix["Disponible"]] != "No" and not f[ix["Precio_Entrada"]].strip() and f[ix["Precio_Entrada_Lista"]].strip():
        f[ix["Precio_Entrada"]] = f[ix["Precio_Entrada_Lista"]]
        n_entrada += 1
        pe, ps = num(f[ix["Precio_Entrada"]]), num(f[ix["Precio_Salida"]])
        if pe and ps:
            f[ix["Margen_Abs"]] = round(ps - pe, 2)
            f[ix["Margen_Pct"]] = round((ps - pe) / ps * 100, 1)

for col in ("Precio_Entrada", "Margen_Abs", "Margen_Pct", "Precio_Entrada_Lista"):
    L = gspread.utils.rowcol_to_a1(1, ix[col] + 1)[:-1]
    ws.update(range_name=f"{L}2:{L}{len(filas) + 1}", values=[[f[ix[col]]] for f in filas],
              value_input_option="USER_ENTERED")
print(f"Precio_Entrada_Lista completado en {n_lista} filas FANNY.")
print(f"Precio_Entrada tomado de la lista en {n_entrada} filas FANNY (márgenes recalculados).")
