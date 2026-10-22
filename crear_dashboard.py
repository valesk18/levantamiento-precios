"""
Crea (o rehace) el DASHBOARD de análisis dentro del Google Sheet Levantamiento_Precios.

Hojas que crea:
  - Base_Dash : copia viva de Registro_Precios + columnas calculadas (zona, categoría, marca, PDV, filtro).
  - Dashboard : filtros (fechas, zona, ruta, categoría), KPIs, tablas y gráficos.
Solo LEE Registro_Precios y los maestros: no modifica datos ni afecta a la app.

Uso (con valesK activo, dentro de precios_app):
    python crear_dashboard.py
"""
import re
import tomllib

import gspread
import pandas as pd

cfg = tomllib.load(open(".streamlit/secrets.toml", "rb"))
sh = gspread.service_account_from_dict(dict(cfg["gcp_service_account"])).open_by_key(cfg["SHEET_ID"])

B = "Base_Dash"
D = "Dashboard"
N_SKU = 10           # filas de la tabla por SKU
N_RUTA = 24          # filas de la tabla por ruta


# ───────────────────────── utilidades ─────────────────────────
def localizar(formula: str, sep: str) -> str:
    """Cambia ',' por ';' fuera de las comillas (para Sheets con configuración regional en español)."""
    if sep == ",":
        return formula
    out, dentro = [], False
    for ch in formula:
        if ch == '"':
            dentro = not dentro
        out.append(";" if (ch == "," and not dentro) else ch)
    return "".join(out)


def hoja(nombre, filas=1000, cols=40):
    """Devuelve la hoja vacía (sin datos, gráficos ni celdas unidas); la crea si no existe."""
    try:
        ws = sh.worksheet(nombre)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(nombre, rows=filas, cols=cols)
    meta = sh.fetch_sheet_metadata()["sheets"]
    graficos = [c["chartId"] for x in meta if x["properties"]["sheetId"] == ws.id for c in x.get("charts", [])]
    reglas = [x for x in meta if x["properties"]["sheetId"] == ws.id][0].get("conditionalFormats", [])
    reqs = [{"deleteEmbeddedObject": {"objectId": g}} for g in graficos]
    reqs += [{"deleteConditionalFormatRule": {"sheetId": ws.id, "index": 0}} for _ in reglas]
    reqs += [{"unmergeCells": {"range": {"sheetId": ws.id}}}]
    if ws.row_count < filas:
        reqs += [{"updateSheetProperties": {"properties": {"sheetId": ws.id, "gridProperties": {"rowCount": filas}},
                                            "fields": "gridProperties.rowCount"}}]
    sh.batch_update({"requests": reqs})
    ws.clear()
    return ws


# ───────────────────────── 1) Base_Dash ─────────────────────────
ENC_EXTRA = ["Zona", "Categoria", "Producto_FANNY", "Marca", "Distribuidor", "PDV_Key", "En_Filtro"]
# Columnas Registro_Precios A..Y ; extras Z..AF
BASE = [
    ("A1", '=ARRAYFORMULA(Registro_Precios!A1:Y)'),
    ("Z1", "Zona"), ("AA1", "Categoria"), ("AB1", "Producto_FANNY"), ("AC1", "Marca"),
    ("AD1", "Distribuidor"), ("AE1", "PDV_Key"), ("AF1", "En_Filtro"),
    ("Z2", '=ARRAYFORMULA(IF(A2:A="",,IF(REGEXMATCH(H2:H,"^(MA_|HORECA|MERCADOS|OFICINA|SUPER)"),"MANAGUA",'
           'IF(REGEXMATCH(H2:H,"^CENTRO"),"CENTRO",IF(REGEXMATCH(H2:H,"^NORTE"),"NORTE","OCCIDENTE/SURORIENTE")))))'),
    ("AA2", '=ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(K2:K,Maestro_SKU!A:E,5,0),"")))'),
    ("AB2", '=ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(K2:K,Maestro_SKU!A:B,2,0),K2:K)))'),
    ("AC2", '=ARRAYFORMULA(IF(A2:A="",,IF(L2:L="Propio","FANNY",IFERROR(VLOOKUP(M2:M,Maestro_Competencia!C:D,2,0),""))))'),
    ("AD2", '=ARRAYFORMULA(IF(A2:A="",,IF(L2:L="Propio","SUPLISA",IFERROR(VLOOKUP(M2:M,Maestro_Competencia!C:E,3,0),""))))'),
    ("AE2", '=ARRAYFORMULA(IF(A2:A="",,F2:F&"|"&TEXT(B2:B+0,"yyyy-mm-dd")))'),
    ("AF2", '=ARRAYFORMULA(IF(A2:A="",,IF((B2:B+0>=Dashboard!C5)*(B2:B+0<=Dashboard!C6)'
            '*((Dashboard!C7="Todas")+(Z2:Z=Dashboard!C7)>0)*((Dashboard!C8="Todas")+(H2:H=Dashboard!C8)>0)'
            '*((Dashboard!C9="Todas")+(AA2:AA=Dashboard!C9)>0),1,0)))'),
    # Listas para los filtros
    ("AH1", "Lista_Zona"), ("AH2", '=VSTACK("Todas","MANAGUA","CENTRO","NORTE","OCCIDENTE/SURORIENTE")'),
    ("AI1", "Lista_Ruta"), ("AI2", '=VSTACK("Todas",SORT(UNIQUE(FILTER(Maestro_Ruta!A2:A,Maestro_Ruta!A2:A<>""))))'),
    ("AJ1", "Lista_Categoria"), ("AJ2", '=VSTACK("Todas",SORT(UNIQUE(FILTER(Maestro_SKU!E2:E,Maestro_SKU!E2:E<>""))))'),
]

# ───────────────────────── 2) Dashboard ─────────────────────────
F = f"{B}!AF:AF,1"         # criterio "dentro del filtro"
R = lambda c: f"{B}!{c}:{c}"


def kpis():
    return [
        ("B12", "PDV levantados"), ("B13", f'=IFERROR(COUNTUNIQUE(FILTER({B}!AE2:AE,{B}!AF2:AF=1)),0)'),
        ("D12", "Rutas activas"), ("D13", f'=IFERROR(COUNTUNIQUE(FILTER({B}!H2:H,{B}!AF2:AF=1)),0)'),
        ("F12", "Municipios"), ("F13", f'=IFERROR(COUNTUNIQUE(FILTER({B}!I2:I,{B}!AF2:AF=1)),0)'),
        ("H12", "Disponibilidad FANNY"),
        ("H13", f'=IFERROR(COUNTIFS({F},{R("L")},"Propio",{R("V")},"Sí")/COUNTIFS({F},{R("L")},"Propio"),"")'),
        ("J12", "Margen PDV FANNY"),
        ("J13", f'=IFERROR(AVERAGEIFS({R("R")},{F},{R("L")},"Propio",{R("R")},"<>")/100,"")'),
        ("L12", "Índice de precio FANNY"),
        ("L13", f'=IFERROR(AVERAGE(FILTER(H19:H{18 + N_SKU},ISNUMBER(H19:H{18 + N_SKU}))),"")'),
        ("N12", "Precios competencia"),
        ("N13", f'=COUNTIFS({F},{R("L")},"Competencia",{R("P")},">0")'),
        ("P12", "Competencia en promoción"),
        ("P13", f'=IFERROR(COUNTIFS({F},{R("L")},"Competencia",{R("W")},"Sí")/COUNTIFS({F},{R("L")},"Competencia",{R("P")},">0"),"")'),
    ]


def tabla_sku():
    enc = ["SKU", "Producto FANNY", "Salida FANNY C$", "Lista zona C$", "C$/UM FANNY", "C$/UM competencia",
           "Índice FANNY", "Lectura", "Margen PDV", "Disponibilidad", "PDV", "Competidor más barato", "C$/UM"]
    celdas = [("B17", "1. POSICIÓN DE PRECIO Y MARGEN POR SKU FANNY  (Índice = C$/UM FANNY ÷ C$/UM competencia × 100; sin promociones)")]
    celdas += [(f"{c}18", t) for c, t in zip("BCDEFGHIJKLMN", enc)]
    for i in range(N_SKU):
        r = 19 + i
        k = f"$B{r}"
        base = f'{F},{R("K")},{k},{R("L")},"Propio"'
        celdas += [
            (f"B{r}", f"=IFERROR(INDEX(Maestro_SKU!A2:A,{i + 1}),\"\")"),
            (f"C{r}", f'=IF({k}="",,IFERROR(VLOOKUP({k},Maestro_SKU!A:B,2,0),""))'),
            (f"D{r}", f'=IF({k}="",,IFERROR(AVERAGEIFS({R("P")},{base},{R("P")},">0"),""))'),
            (f"E{r}", f'=IF({k}="",,IFERROR(AVERAGEIFS({R("Y")},{base},{R("Y")},">0"),""))'),
            (f"F{r}", f'=IF({k}="",,IFERROR(AVERAGEIFS({R("U")},{base},{R("U")},">0"),""))'),
            (f"G{r}", f'=IF({k}="",,IFERROR(AVERAGEIFS({R("U")},{F},{R("K")},{k},{R("L")},"Competencia",{R("W")},"No",{R("U")},">0"),""))'),
            (f"H{r}", f'=IF(OR(F{r}="",G{r}=""),"",ROUND(F{r}/G{r}*100,0))'),
            (f"I{r}", f'=IF(H{r}="","",IF(H{r}>105,"FANNY más cara",IF(H{r}<95,"FANNY más barata","Paridad")))'),
            (f"J{r}", f'=IF({k}="",,IFERROR(AVERAGEIFS({R("R")},{base},{R("R")},"<>")/100,""))'),
            (f"K{r}", f'=IF({k}="",,IFERROR(COUNTIFS({base},{R("V")},"Sí")/COUNTIFS({base}),""))'),
            (f"L{r}", f'=IF({k}="",,COUNTIFS({base}))'),
            (f"M{r}", f'=IF({k}="",,IFERROR(QUERY({B}!A:AF,"select AC, avg(U) where K=\'"&{k}&"\' and L=\'Competencia\' '
                      f'and W=\'No\' and U>0 and AF=1 group by AC order by avg(U) asc limit 1 label AC \'\', avg(U) \'\'",0),""))'),
        ]
    return celdas


def tabla_ruta():
    r0 = 19 + N_SKU + 3                      # fila de título
    enc = ["Ruta", "Zona", "PDV", "Disponibilidad FANNY", "Margen PDV FANNY", "Precios competencia", "Último registro"]
    celdas = [(f"B{r0}", "2. AVANCE Y DESEMPEÑO POR RUTA")]
    celdas += [(f"{c}{r0 + 1}", t) for c, t in zip("BCDEFGH", enc)]
    a, z = r0 + 2, r0 + 1 + N_RUTA
    rr = f"B{a}:B{z}"
    celdas += [
        (f"B{a}", f'=IFERROR(SORT(UNIQUE(FILTER({B}!H2:H,{B}!AF2:AF=1))),"Sin datos")'),
        (f"C{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,IFERROR(VLOOKUP(x,{B}!H:Z,19,0),""))))'),
        (f"D{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,COUNTUNIQUEIFS({B}!AE:AE,{B}!H:H,x,{B}!AF:AF,1))))'),
        (f"E{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,IFERROR(COUNTIFS({B}!H:H,x,{F},{R("L")},"Propio",{R("V")},"Sí")'
                  f'/COUNTIFS({B}!H:H,x,{F},{R("L")},"Propio"),""))))'),
        (f"F{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,IFERROR(AVERAGEIFS({R("R")},{B}!H:H,x,{F},{R("L")},"Propio",{R("R")},"<>")/100,""))))'),
        (f"G{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,COUNTIFS({B}!H:H,x,{F},{R("L")},"Competencia",{R("P")},">0"))))'),
        (f"H{a}", f'=BYROW({rr},LAMBDA(x,IF(x="",,IFERROR(MAXIFS({B}!C:C,{B}!H:H,x,{B}!AF:AF,1),""))))'),
    ]
    return celdas, r0, a, z


def tablas_derecha(r0):
    c = []
    # 3. Salida FANNY por zona (pivote)
    c += [("J" + str(r0), "3. PRECIO DE SALIDA FANNY PROMEDIO POR ZONA (C$)")]
    c += [("J" + str(r0 + 1), f'=IFERROR(QUERY({B}!A:AF,"select AB, avg(P) where L=\'Propio\' and P>0 and AF=1 '
                              f'group by AB pivot Z label AB \'Producto\'",1),"Sin datos")')]
    # 4. Presencia de competidores
    r4 = r0 + 15
    c += [("J" + str(r4), "4. PRESENCIA DE COMPETIDORES (veces encontrados con existencia)")]
    c += [("J" + str(r4 + 1), f'=IFERROR(QUERY({B}!A:AF,"select AC, AD, count(A), avg(U) where L=\'Competencia\' '
                              f'and V=\'Sí\' and AF=1 group by AC, AD order by count(A) desc limit 12 '
                              f'label AC \'Marca\', AD \'Distribuidor\', count(A) \'Presencia\', avg(U) \'C$/UM prom\'",1),"Sin datos")')]
    return c, r4


# ───────────────────────── construir ─────────────────────────
wb = hoja(B, 20000, 40)
wd = hoja(D, 120, 18)

sep = ","
for intento in (",", ";"):
    sep = intento
    wb.update(range_name="A1", values=[[""]])
    for celda, v in BASE:
        wb.update(range_name=celda, values=[[localizar(v, sep)]], value_input_option="USER_ENTERED")
    prueba = wb.get_values("Z1:Z2") + wb.get_values("AH2:AH2")
    if not any(str(x).startswith("#") for fila in prueba for x in fila):
        break

# Dashboard: encabezado y filtros
celdas = [
    ("B2", "DASHBOARD · LEVANTAMIENTO DE PRECIOS FANNY vs. COMPETENCIA"),
    ("B3", "Fuente: Registro_Precios (se actualiza solo con cada guardado de la app). Cambie los filtros en amarillo."),
    ("B5", "Desde"), ("C5", f'=IFERROR(MIN(FILTER({B}!B2:B+0,{B}!B2:B<>"")),TODAY())'),
    ("B6", "Hasta"), ("C6", "=TODAY()"),
    ("B7", "Zona"), ("C7", "Todas"),
    ("B8", "Ruta"), ("C8", "Todas"),
    ("B9", "Categoría"), ("C9", "Todas"),
    ("E5", "Cómo leer el índice:"), ("E6", "> 105  FANNY más cara que la competencia"),
    ("E7", "95–105  Paridad"), ("E8", "< 95  FANNY más barata"),
    ("E9", "Margen PDV = (Salida − Entrada) ÷ Salida · si no hay entrada observada se usa el precio de lista de la zona"),
]
celdas += kpis() + tabla_sku()
c_ruta, r0, a, z = tabla_ruta()
c_der, r4 = tablas_derecha(r0)
celdas += c_ruta + c_der
wd.batch_update([{"range": c, "values": [[localizar(v, sep) if str(v).startswith("=") else v]]} for c, v in celdas],
                value_input_option="USER_ENTERED")

# ───────────────────────── formato ─────────────────────────
AZUL = {"red": 0.071, "green": 0.196, "blue": 0.29}
GRIS = {"red": 0.93, "green": 0.95, "blue": 0.97}
AMAR = {"red": 1, "green": 0.95, "blue": 0.8}
BLANCO = {"red": 1, "green": 1, "blue": 1}
fmt = [
    ("B2", {"textFormat": {"bold": True, "fontSize": 16, "foregroundColor": AZUL}}),
    ("B3", {"textFormat": {"italic": True, "foregroundColor": {"red": .4, "green": .4, "blue": .4}}}),
    ("B5:B9", {"textFormat": {"bold": True}}),
    ("C5:C9", {"backgroundColor": AMAR, "textFormat": {"bold": True}}),
    ("C5:C6", {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}}),
    ("E5", {"textFormat": {"bold": True}}),
    ("B12:Q12", {"backgroundColor": AZUL, "textFormat": {"bold": True, "foregroundColor": BLANCO, "fontSize": 9},
                 "wrapStrategy": "WRAP", "horizontalAlignment": "CENTER"}),
    ("B13:Q13", {"backgroundColor": GRIS, "textFormat": {"bold": True, "fontSize": 18, "foregroundColor": AZUL},
                 "horizontalAlignment": "CENTER"}),
    ("H13", {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}),
    ("J13", {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}),
    ("L13", {"numberFormat": {"type": "NUMBER", "pattern": "0"}}),
    ("P13", {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}),
    (f"B17", {"textFormat": {"bold": True, "fontSize": 12, "foregroundColor": AZUL}}),
    ("B18:N18", {"backgroundColor": AZUL, "textFormat": {"bold": True, "foregroundColor": BLANCO},
                 "wrapStrategy": "WRAP"}),
    (f"D19:G{18 + N_SKU}", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}),
    (f"F19:G{18 + N_SKU}", {"numberFormat": {"type": "NUMBER", "pattern": "0.0000"}}),
    (f"J19:K{18 + N_SKU}", {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}),
    (f"N19:N{18 + N_SKU}", {"numberFormat": {"type": "NUMBER", "pattern": "0.0000"}}),
    (f"B{r0}", {"textFormat": {"bold": True, "fontSize": 12, "foregroundColor": AZUL}}),
    (f"B{r0 + 1}:H{r0 + 1}", {"backgroundColor": AZUL, "textFormat": {"bold": True, "foregroundColor": BLANCO},
                               "wrapStrategy": "WRAP"}),
    (f"E{a}:F{z}", {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}),
    (f"H{a}:H{z}", {"numberFormat": {"type": "DATE_TIME", "pattern": "dd/mm hh:mm"}}),
    (f"J{r0}", {"textFormat": {"bold": True, "fontSize": 12, "foregroundColor": AZUL}}),
    (f"J{r0 + 1}:Q{r0 + 1}", {"backgroundColor": GRIS, "textFormat": {"bold": True}}),
    (f"K{r0 + 2}:Q{r0 + 12}", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}),
    (f"J{r4}", {"textFormat": {"bold": True, "fontSize": 12, "foregroundColor": AZUL}}),
    (f"J{r4 + 1}:M{r4 + 1}", {"backgroundColor": GRIS, "textFormat": {"bold": True}}),
    (f"M{r4 + 2}:M{r4 + 14}", {"numberFormat": {"type": "NUMBER", "pattern": "0.0000"}}),
]
wd.batch_format([{"range": r, "format": f} for r, f in fmt])

reqs = [
    # anchos de columna
    {"updateDimensionProperties": {"range": {"sheetId": wd.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                                   "properties": {"pixelSize": 20}, "fields": "pixelSize"}},
    {"updateDimensionProperties": {"range": {"sheetId": wd.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 17},
                                   "properties": {"pixelSize": 118}, "fields": "pixelSize"}},
    {"updateDimensionProperties": {"range": {"sheetId": wd.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
                                   "properties": {"pixelSize": 190}, "fields": "pixelSize"}},
    {"updateSheetProperties": {"properties": {"sheetId": wd.id, "gridProperties": {"hideGridlines": True}},
                               "fields": "gridProperties.hideGridlines"}},
    # unir celdas de cada KPI (etiqueta y valor ocupan 2 columnas)
    *[{"mergeCells": {"range": {"sheetId": wd.id, "startRowIndex": fila, "endRowIndex": fila + 1,
                                "startColumnIndex": col, "endColumnIndex": col + 2}, "mergeType": "MERGE_ALL"}}
      for fila in (11, 12) for col in range(1, 17, 2)],
    # listas desplegables de filtros
    *[{"setDataValidation": {"range": {"sheetId": wd.id, "startRowIndex": fila, "endRowIndex": fila + 1,
                                       "startColumnIndex": 2, "endColumnIndex": 3},
                             "rule": {"condition": {"type": "ONE_OF_RANGE",
                                                    "values": [{"userEnteredValue": f"={B}!{col}2:{col}60"}]},
                                      "showCustomUi": True, "strict": True}}}
      for fila, col in ((6, "AH"), (7, "AI"), (8, "AJ"))],
    *[{"setDataValidation": {"range": {"sheetId": wd.id, "startRowIndex": fila, "endRowIndex": fila + 1,
                                       "startColumnIndex": 2, "endColumnIndex": 3},
                             "rule": {"condition": {"type": "DATE_IS_VALID"}, "strict": True}}}
      for fila in (4, 5)],
    # semáforo del índice
    {"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": wd.id, "startRowIndex": 18, "endRowIndex": 18 + N_SKU, "startColumnIndex": 7, "endColumnIndex": 9}],
        "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": localizar("=AND(ISNUMBER($H19),$H19>105)", sep)}]},
                        "format": {"backgroundColor": {"red": .99, "green": .85, "blue": .82}}}}}},
    {"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": wd.id, "startRowIndex": 18, "endRowIndex": 18 + N_SKU, "startColumnIndex": 7, "endColumnIndex": 9}],
        "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": localizar("=AND(ISNUMBER($H19),$H19<95)", sep)}]},
                        "format": {"backgroundColor": {"red": .85, "green": .95, "blue": .87}}}}}},
    # disponibilidad baja (< 80%): por ruta (col E) y por SKU (col K)
    *[{"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": wd.id, "startRowIndex": f0 - 1, "endRowIndex": f1, "startColumnIndex": ci, "endColumnIndex": ci + 1}],
        "booleanRule": {"condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": localizar(f"=AND(ISNUMBER({cl}{f0}),{cl}{f0}<80/100)", sep)}]},
                        "format": {"textFormat": {"foregroundColor": {"red": .75, "green": .1, "blue": .1}, "bold": True}}}}}}
      for f0, f1, ci, cl in ((a, z, 4, "E"), (19, 18 + N_SKU, 10, "K"))],
]


def grafico(titulo, tipo, dom, series, fila, col, ancho=520, alto=300):
    rng = lambda c1, c2, f1, f2: {"sources": [{"sheetId": wd.id, "startRowIndex": f1, "endRowIndex": f2,
                                               "startColumnIndex": c1, "endColumnIndex": c2}]}
    return {"addChart": {"chart": {
        "spec": {"title": titulo, "basicChart": {
            "chartType": tipo, "legendPosition": "NO_LEGEND", "headerCount": 1,
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": rng(*dom)}}],
            "series": [{"series": {"sourceRange": rng(*s)}, "targetAxis": "LEFT_AXIS" if tipo == "COLUMN" else "BOTTOM_AXIS",
                        "color": AZUL} for s in series]}},
        "position": {"overlayPosition": {"anchorCell": {"sheetId": wd.id, "rowIndex": fila, "columnIndex": col},
                                         "widthPixels": ancho, "heightPixels": alto}}}}}


fin = z + 3
reqs += [
    grafico("Índice de precio FANNY por SKU (100 = paridad)", "BAR",
            (2, 3, 17, 18 + N_SKU), [(7, 8, 17, 18 + N_SKU)], fin, 1),
    grafico("PDV levantados por ruta", "COLUMN", (1, 2, r0, z), [(3, 4, r0, z)], fin, 6),
    grafico("Disponibilidad FANNY por SKU", "BAR", (2, 3, 17, 18 + N_SKU), [(10, 11, 17, 18 + N_SKU)], fin + 17, 1),
    grafico("Margen PDV FANNY por SKU", "BAR", (2, 3, 17, 18 + N_SKU), [(9, 10, 17, 18 + N_SKU)], fin + 17, 6),
]
sh.batch_update({"requests": reqs})

# Ubicar el Dashboard como primera hoja y ocultar Base_Dash
sh.batch_update({"requests": [
    {"updateSheetProperties": {"properties": {"sheetId": wd.id, "index": 0}, "fields": "index"}},
    {"updateSheetProperties": {"properties": {"sheetId": wb.id, "hidden": True}, "fields": "hidden"}},
]})

# ───────────────────────── verificación ─────────────────────────
vals = wd.get_values("B12:Q13")
err = [f"{gspread.utils.rowcol_to_a1(i, j + 1)}={x}"
       for i, fila in enumerate(wd.get_values("A1:Q120"), 1) for j, x in enumerate(fila) if str(x).startswith("#")]
print(f"Dashboard creado (separador de fórmulas '{sep}').")
if len(vals) > 1:
    for et, va in zip(vals[0][::2], vals[1][::2]):
        print(f"   {et}: {va}")
print("Celdas con error:", err[:15] if err else "ninguna")

# Copia local para revisión del análisis
v = sh.worksheet("Registro_Precios").get_all_values()
pd.DataFrame(v[1:], columns=v[0]).to_csv("data/registro_export.csv", index=False, encoding="utf-8")
print("Copia para revisión: data/registro_export.csv")
