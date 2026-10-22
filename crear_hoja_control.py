"""
Crea (o rehace) la hoja "Control" en el Google Sheet Levantamiento_Precios.
Solo LEE Registro_Precios con fórmulas: no lo modifica ni afecta a la app.

Uso (con valesK activo, dentro de precios_app):
    python crear_hoja_control.py
"""
import tomllib

import gspread

cfg = tomllib.load(open(".streamlit/secrets.toml", "rb"))
sh = gspread.service_account_from_dict(dict(cfg["gcp_service_account"])).open_by_key(cfg["SHEET_ID"])
R = "Registro_Precios"


def contenido(sep: str) -> list[list[str]]:
    f = (lambda s: s.replace(",", ";")) if sep == ";" else (lambda s: s)
    return [
        ["CONTROL DEL LEVANTAMIENTO", "", "", "", "Solo lectura: no modifica Registro_Precios"],
        [],
        ["Filas con datos", f(f"=COUNTA({R}!A2:A)")],
        ["Levantamientos (PDV guardados)", f(f"=COUNTUNIQUE({R}!A2:A)")],
        ["Rutas que han registrado", f(f"=COUNTUNIQUE({R}!H2:H)")],
        ["Último guardado", f(f"=MAX({R}!C2:C)")],
        ["Levantamientos de hoy", f(f"=IFERROR(COUNTUNIQUE(FILTER({R}!A2:A,{R}!B2:B=TODAY())),0)")],
        [],
        ["POR FECHA Y RUTA", "", "", "", "MUNICIPIOS REGISTRADOS (revisar nombres raros)"],
        [f(f"=QUERY({R}!A:H,\"select B, H, count(A) where A<>'' group by B, H "
             f"label B 'Fecha', H 'Ruta', count(A) 'Filas'\",1)"),
         "", "", "",
         f(f'=SORT(UNIQUE(FILTER({R}!I2:J,{R}!I2:I<>"")))')],
    ]


try:
    ws = sh.worksheet("Control")
except gspread.WorksheetNotFound:
    ws = sh.add_worksheet("Control", rows=300, cols=8)

# Probar con ',' y, si Google no entiende las fórmulas (configuración regional), con ';'
for sep in (",", ";"):
    ws.clear()
    ws.update(range_name="A1", values=contenido(sep), value_input_option="USER_ENTERED")
    v = ws.get_values("B3:B7")
    if not any(str(c).startswith("#") for fila in v for c in fila):
        break

ws.format("A1", {"textFormat": {"bold": True, "fontSize": 12}})
ws.format("A3:A7", {"textFormat": {"bold": True}})
ws.format("A9:E9", {"textFormat": {"bold": True}})
ws.format("B6", {"numberFormat": {"type": "DATE_TIME", "pattern": "dd/mm/yyyy hh:mm"}})
ws.format("A10:A300", {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}})

print(f"Hoja 'Control' lista (separador de fórmulas '{sep}').")
for fila in ws.get_values("A3:B7"):
    print("  ", fila[0], "→", fila[1] if len(fila) > 1 else "")
