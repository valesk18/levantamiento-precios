"""
Levantamiento de Precios · Propio vs. Competencia (máx. 3 competidores por SKU)
-------------------------------------------------------------------------------
Modo LOCAL  : sin secrets -> lee/escribe CSV en ./data (para probar en VS Code).
Modo SHEETS : con [gcp_service_account] y SHEET_ID en .streamlit/secrets.toml
              -> lee maestros y escribe en la hoja Registro_Precios de Google Sheets.
Ejecutar    : streamlit run app.py
"""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import hashlib
import time
import hmac
import unicodedata
import uuid

import pandas as pd
import streamlit as st

from divipola import MUNICIPIOS

MAX_COMP = 3
DATA_DIR = Path(__file__).parent / "data"
TZ_OFFSET_H = -6  # Nicaragua (UTC-6), para la hora de registro en la nube
MAX_INTENTOS = 5  # intentos de PIN por sesión antes de bloquear
CANALES = ["Detalle", "Horeca", "Mayorista", "Oficina", "Super"]

COLS_REGISTRO = [
    "ID_Lev", "Fecha", "Hora_Registro", "Usuario", "Cliente_ID", "Cliente_Nombre", "Canal", "Ruta",
    "Municipio", "Departamento",
    "SKU_ID", "Tipo", "Producto_ID", "Comp_Orden", "Precio_Entrada", "Precio_Salida",
    "Margen_Abs", "Margen_Pct", "Contenido", "UM", "Precio_UM", "Disponible", "Promo", "Sell_Out",
    "Precio_Entrada_Lista",   # precio SUPLISA por zona (Maestro_SKU) = entrada del PDV para FANNY
]

st.set_page_config(page_title="Levantamiento de Precios", page_icon="🏷️", layout="centered")


# ───────────────────────── 1. Capa de datos ─────────────────────────
def usar_sheets() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "SHEET_ID" in st.secrets
    except Exception:  # no existe secrets.toml
        return False


def canal_de_ruta(fila) -> str:
    """Canal por defecto: columna Canal de Maestro_Ruta; si está vacía, se deduce del código de ruta."""
    c = str(fila.get("Canal", "") or "").strip().title()
    if c in CANALES:
        return c
    ruta = str(fila.get("Ruta", "")).upper()
    for prefijo, canal in [("MA_", "Detalle"), ("HORECA", "Horeca"), ("OFICINA", "Oficina"), ("SUPER", "Super"),
                           ("MERCADOS", "Mayorista"), ("NORTE", "Mayorista"), ("CENTRO", "Mayorista"),
                           ("OCCIDENTE", "Mayorista"), ("SURORIENTE", "Mayorista")]:
        if ruta.startswith(prefijo):
            return canal
    return ""


def hash_pin(pin: str) -> str:
    """SHA-256 de (PIN_SALT + PIN). La sal vive en secrets, no en el Sheet."""
    try:
        sal = st.secrets.get("PIN_SALT", "")
    except Exception:
        sal = ""
    return hashlib.sha256((sal + str(pin).strip()).encode("utf-8")).hexdigest()


@st.cache_resource
def libro_sheets():
    import gspread
    gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    return gc.open_by_key(st.secrets["SHEET_ID"])


def con_reintentos(fn, intentos: int = 4):
    """Reintenta llamadas a Google si responde 'cuota excedida' (429) o falla temporal (5xx)."""
    import gspread
    for i in range(intentos):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            codigo = getattr(getattr(e, "response", None), "status_code", None)
            if codigo in (429, 500, 502, 503) and i < intentos - 1:
                time.sleep(2 * 2 ** i)          # 2, 4, 8 segundos
                continue
            raise


@st.cache_data(ttl=600, show_spinner="Cargando maestros…")
def leer_tabla(nombre: str) -> pd.DataFrame:
    if usar_sheets():
        df = pd.DataFrame(con_reintentos(lambda: libro_sheets().worksheet(nombre).get_all_records()))
    else:
        df = pd.read_csv(DATA_DIR / f"{nombre}.csv", dtype=str)
    return df.fillna("").astype(str).apply(lambda c: c.str.strip())


def leer_registro() -> pd.DataFrame:
    """Sin caché: se usa para validar duplicados con lo último guardado."""
    if usar_sheets():
        df = pd.DataFrame(libro_sheets().worksheet("Registro_Precios").get_all_records())
    else:
        ruta = DATA_DIR / "Registro_Precios.csv"
        df = pd.read_csv(ruta, dtype=str) if ruta.exists() else pd.DataFrame()
    # Tolera registros con estructura anterior (columnas faltantes quedan vacías)
    if "Precio" in df and "Precio_Salida" not in df:
        df = df.rename(columns={"Precio": "Precio_Salida"})
    return df.reindex(columns=COLS_REGISTRO).fillna("")


@st.cache_data(ttl=120, show_spinner=False)
def claves_registro() -> pd.DataFrame:
    """Solo Fecha y Cliente_Nombre, en UNA lectura y con caché de 2 min.
    Antes se leía todo el registro en cada clic y se agotaba la cuota de lectura de Google."""
    if usar_sheets():
        ws = libro_sheets().worksheet("Registro_Precios")
        c_f = COLS_REGISTRO.index("Fecha") + 1
        c_c = COLS_REGISTRO.index("Cliente_Nombre") + 1
        rangos = [f"{gspread_col(c_f)}2:{gspread_col(c_f)}", f"{gspread_col(c_c)}2:{gspread_col(c_c)}"]
        fechas, clientes = con_reintentos(lambda: ws.batch_get(rangos))
        fechas = [r[0] if r else "" for r in fechas]
        clientes = [r[0] if r else "" for r in clientes]
        n = max(len(fechas), len(clientes))
        fechas += [""] * (n - len(fechas)); clientes += [""] * (n - len(clientes))
        return pd.DataFrame({"Fecha": fechas, "Cliente_Nombre": clientes})
    return leer_registro()[["Fecha", "Cliente_Nombre"]]


def gspread_col(n: int) -> str:
    """1 -> A, 6 -> F, 27 -> AA."""
    letras = ""
    while n:
        n, r = divmod(n - 1, 26)
        letras = chr(65 + r) + letras
    return letras


def guardar_filas(filas: list[dict]) -> None:
    nuevas = pd.DataFrame(filas).reindex(columns=COLS_REGISTRO).fillna("")
    if usar_sheets():
        ws = libro_sheets().worksheet("Registro_Precios")
        encabezado = con_reintentos(lambda: ws.row_values(1))
        if encabezado != COLS_REGISTRO:
            if len(ws.get_all_values()) <= 1:      # hoja vacía: se corrige el encabezado
                ws.update(range_name="A1", values=[COLS_REGISTRO])
            else:
                raise RuntimeError("Los encabezados de Registro_Precios no coinciden con la app. "
                                   "Deben ser, en este orden: " + ", ".join(COLS_REGISTRO))
        # INSERT_ROWS: inserta filas nuevas al final; nunca sobrescribe filas existentes
        # (aunque la hoja tenga filtros o un rango de tabla que termine antes de los datos)
        con_reintentos(lambda: ws.append_rows(nuevas.values.tolist(), value_input_option="USER_ENTERED",
                                              insert_data_option="INSERT_ROWS", table_range="A1"))
    else:
        ruta = DATA_DIR / "Registro_Precios.csv"
        if ruta.exists():                           # migra un CSV con estructura anterior
            previo = pd.read_csv(ruta, dtype=str)
            if "Precio" in previo and "Precio_Salida" not in previo:
                previo = previo.rename(columns={"Precio": "Precio_Salida"})
            if list(previo.columns) != COLS_REGISTRO:
                previo.reindex(columns=COLS_REGISTRO).to_csv(ruta, index=False, encoding="utf-8")
        nuevas.to_csv(ruta, mode="a", index=False, header=not ruta.exists(), encoding="utf-8")


def activos(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Activo"].str.upper().isin(["SI", "SÍ", "1", "TRUE"])] if "Activo" in df else df


# ───────────────────────── 2. Maestros ─────────────────────────
try:
    m_sku = activos(leer_tabla("Maestro_SKU"))
    m_comp = activos(leer_tabla("Maestro_Competencia"))
    m_ruta = leer_tabla("Maestro_Ruta")      # acceso por RUTA: columnas Ruta, Descripcion, PIN_Hash
    try:                                      # opcional: Ruta, Municipio
        m_muni = leer_tabla("Maestro_Municipio")
    except Exception:
        m_muni = pd.DataFrame(columns=["Ruta", "Municipio", "Departamento"])
except Exception as e:
    st.error(f"No se pudieron leer los maestros. Revise nombres de hojas/columnas. Detalle: {e}")
    st.stop()

# Regla de negocio: máximo 3 competidores por SKU, ordenados por Comp_Orden
m_comp["Comp_Orden"] = pd.to_numeric(m_comp["Comp_Orden"], errors="coerce")
m_comp = m_comp.sort_values(["SKU_ID", "Comp_Orden"]).groupby("SKU_ID").head(MAX_COMP)

if "form_id" not in st.session_state:
    st.session_state.form_id = 0      # cambia tras guardar -> limpia el formulario
fid = st.session_state.form_id


def num(x) -> float | None:
    try:
        v = float(str(x).replace(",", "."))
        return v if v > 0 else None
    except ValueError:
        return None


# ───────────────────────── 3. Encabezado y PDV ─────────────────────────
st.title("🏷️ Levantamiento de Precios")
st.caption(f"Propio vs. hasta {MAX_COMP} competidores por SKU · índice por unidad de medida · "
           f"{'Google Sheets' if usar_sheets() else 'MODO LOCAL (CSV)'}")

if msg := st.session_state.pop("ok_msg", None):
    st.success(msg)

# ── Acceso con PIN ──
# Maestro_Ruta guarda PIN_Hash (nunca el PIN en texto). Genere hashes con generar_pin.py
desc = m_ruta["Descripcion"] if "Descripcion" in m_ruta else pd.Series("", index=m_ruta.index)
m_ruta["Etiqueta"] = m_ruta["Ruta"] + desc.map(lambda d: f" · {d}" if d else "")
if "ruta_sesion" not in st.session_state:
    st.session_state.intentos = st.session_state.get("intentos", 0)
    if st.session_state.intentos >= MAX_INTENTOS:
        st.error("Demasiados intentos fallidos. Cierre la app y comuníquese con Inteligencia Comercial.")
        st.stop()
    with st.form("login"):
        etiqueta = st.selectbox("Ruta", m_ruta["Etiqueta"].tolist(), index=None,
                                placeholder="Seleccione su ruta")
        pin = st.text_input("PIN de la ruta", type="password", max_chars=6)
        entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)
    if entrar:
        fila = m_ruta[m_ruta["Etiqueta"] == etiqueta]
        if etiqueta and not fila.empty and hmac.compare_digest(hash_pin(pin), fila.iloc[0]["PIN_Hash"]):
            st.session_state.ruta_sesion = fila.iloc[0]["Ruta"]
            st.session_state.intentos = 0
            st.rerun()
        st.session_state.intentos += 1
        st.error(f"Ruta o PIN incorrecto. Intentos restantes: "
                 f"{MAX_INTENTOS - st.session_state.intentos}")
    st.stop()

fila_ruta = m_ruta[m_ruta["Ruta"] == st.session_state.ruta_sesion].iloc[0]
vend = {"Ruta": fila_ruta["Ruta"], "Usuario": fila_ruta["Ruta"]}
# Zona de precio de la ruta → columna Precio_<ZONA> en Maestro_SKU
ZONAS_PRECIO = [("MA_", "MANAGUA"), ("HORECA", "MANAGUA"), ("MERCADOS", "MANAGUA"), ("OFICINA", "MANAGUA"),
                ("SUPER", "MANAGUA"), ("CENTRO", "CENTRO"), ("NORTE", "NORTE"),
                ("OCCIDENTE", "OCCIDENTE_SURORIENTE"), ("SURORIENTE", "OCCIDENTE_SURORIENTE")]


def zona_precio(fila) -> str:
    """Columna Zona_Precio de Maestro_Ruta si existe; si no, se deduce del código de ruta."""
    z = str(fila.get("Zona_Precio", "") or "").strip().upper().replace(" ", "_")
    if z:
        return z
    ruta = str(fila.get("Ruta", "")).upper()
    return next((zona for pre, zona in ZONAS_PRECIO if ruta.startswith(pre)), "")


zona_ruta = zona_precio(fila_ruta)
# Palabras clave por zona: tolera encabezados combinados (Precio_NORTE_CENTRO) o con errores (OOCIDENTE)
CLAVES_ZONA = {"MANAGUA": ["MANAGUA"], "CENTRO": ["CENTRO"], "NORTE": ["NORTE"],
               "OCCIDENTE_SURORIENTE": ["CIDENTE", "SURORIENTE", "SUR_ORIENTE"]}


def columna_zona(columnas, zona: str) -> str:
    """Primera columna Precio_* de Maestro_SKU (distinta de Precio_Sugerido) que corresponde a la zona."""
    claves = CLAVES_ZONA.get(zona, [zona])
    for c in columnas:
        cu = str(c).upper().replace(" ", "_")
        if cu.startswith("PRECIO_") and cu != "PRECIO_SUGERIDO" and any(k in cu for k in claves):
            return c
    return ""


COL_PRECIO_ZONA = columna_zona(m_sku.columns, zona_ruta)


def precio_lista(s):
    """Precio SUPLISA de la zona (entrada del PDV para el SKU propio). Vacío si no está cargado."""
    return num(str(s.get(COL_PRECIO_ZONA, "") or "").strip()) if COL_PRECIO_ZONA else None


vend_nombre = fila_ruta["Etiqueta"]
c1, c2 = st.columns([2, 1])
c1.markdown(f"🚚 **Ruta {vend_nombre}**")
if c2.button("Salir", use_container_width=True):
    st.session_state.clear()
    st.rerun()
hoy_ni = datetime.now(timezone(timedelta(hours=TZ_OFFSET_H))).date()   # fecha de Nicaragua (el servidor está en UTC)
fecha = hoy_ni                                  # fija: solo se registra el día en curso
st.markdown(f"📅 **Fecha:** {fecha:%d/%m/%Y}")
# ── Cliente digitado (temporal, mientras no exista Maestro_PDV) ──
def normalizar(txt: str) -> str:
    """Mayúsculas, sin tildes ni espacios dobles: 'Pulpería  la Esperanza' -> 'PULPERIA LA ESPERANZA'."""
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return " ".join(t.upper().split())


try:
    previos = claves_registro()
except Exception:
    previos = pd.DataFrame(columns=["Fecha", "Cliente_Nombre"])   # si Google falla, no bloquea la captura
# Lo guardado en esta sesión cuenta como previo aunque la caché aún no lo refleje
previos = pd.concat([previos, pd.DataFrame(st.session_state.get("guardados", []),
                                           columns=["Fecha", "Cliente_Nombre"])], ignore_index=True)

# ── Municipio: solo los de la ruta en sesión (Maestro_Municipio); "Otro" permite escribirlo ──
OTRO = "Otro (escribir)"
DEPARTAMENTOS = ["Boaco", "Carazo", "Chinandega", "Chontales", "Estelí", "Granada", "Jinotega",
                 "León", "Madriz", "Managua", "Masaya", "Matagalpa", "Nueva Segovia", "RACCN",
                 "RACCS", "Río San Juan", "Rivas"]
de_ruta = m_muni[m_muni["Ruta"] == vend["Ruta"]] if not m_muni.empty else m_muni
if "Departamento" not in de_ruta:
    de_ruta = de_ruta.assign(Departamento="")
de_ruta = de_ruta[de_ruta["Municipio"] != ""].assign(
    _n=pd.to_numeric(de_ruta.get("Clientes_BD", 0), errors="coerce")).fillna({"_n": 0})

# 1) Departamento: los 17 del país; primero los que atiende la ruta (según RUTAS_MUN)
deptos_ruta = sorted(de_ruta.loc[de_ruta["Departamento"] != "", "Departamento"].unique(), key=normalizar)
todos_deptos = deptos_ruta + sorted([d for d in MUNICIPIOS if d not in deptos_ruta], key=normalizar)
d1, d2 = st.columns(2)
departamento = d1.selectbox("Departamento", todos_deptos,
                            index=0 if len(deptos_ruta) == 1 else None,   # ruta de un solo depto: ya seleccionado
                            placeholder="Seleccione el departamento", key=f"dep_{fid}") or ""

# 2) Municipio: todos los del departamento (división política); primero los que atiende la ruta
munis_ruta = sorted(de_ruta.loc[de_ruta["Departamento"] == departamento, "Municipio"].unique(), key=normalizar)
munis = (munis_ruta + sorted([m for m in MUNICIPIOS.get(departamento, []) if m not in munis_ruta],
                             key=normalizar)) if departamento else []
sel = d2.selectbox("Municipio", munis + [OTRO] if departamento else [], index=None,
                   placeholder="Seleccione el municipio" if departamento else "Primero el departamento",
                   disabled=not departamento, key=f"muni_{fid}")
if sel == OTRO:
    municipio = st.text_input(f"Escriba el municipio de {departamento}", key=f"munix_{fid}")
else:
    municipio = sel or ""
municipio = normalizar(municipio)
departamento = normalizar(departamento)

c1, c2 = st.columns([2, 1])
cliente = c1.text_input("Nombre del cliente", key=f"cli_{fid}", max_chars=80,
                        placeholder="Ej.: Pulpería La Esperanza")
canal_ruta = canal_de_ruta(fila_ruta)          # canal predominante de la ruta, editable si el PDV es de otro
canal = c2.selectbox("Canal", CANALES, index=CANALES.index(canal_ruta) if canal_ruta in CANALES else None,
                     placeholder="Canal", key=f"canal_{fid}")
cliente = normalizar(cliente)
if len(municipio) < 3 or len(cliente) < 3 or not canal:
    st.info("Complete municipio, nombre del cliente (mínimo 3 letras) y canal.")
    st.stop()
pdv = {"Cliente_ID": "", "Nombre": cliente, "Canal": canal, "Ruta": vend["Ruta"],
       "Municipio": municipio, "Departamento": departamento}
st.caption(f"Cliente **{cliente}** · {municipio} ({departamento or 's/d'}) · "
           f"Canal **{canal}** · Ruta **{vend['Ruta']}**")

st.divider()

# ───────────────────────── 4. Captura por SKU ─────────────────────────
capturas = []  # una entrada por producto observado


def fila_producto(sku_id, tipo, prod_id, nombre, cont, um, orden, extra="", lista=None):
    """Una fila por producto: Precio entrada (lo que paga el PDV, opcional) y salida (al consumidor)."""
    k = f"{fid}_{sku_id}_{prod_id}"
    st.markdown(f"{'**🟦 PROPIO** · ' if tipo == 'Propio' else ''}{nombre}  \n"
                f"<small>{cont} {um}{extra}</small>", unsafe_allow_html=True)
    a, b = st.columns(2)
    entrada = a.text_input("Entrada C$ (opcional)", key=f"e_{k}",
                           placeholder=f"Lista zona: {lista:,.2f}" if lista else "Si se consigue")
    salida = b.text_input("Salida C$", key=f"p_{k}", placeholder="Precio al consumidor")
    f1, f2 = st.columns(2)
    sin_exist = f1.checkbox("Sin existencia", key=f"nd_{k}")
    promo = f2.checkbox("En promoción", key=f"pr_{k}")
    pe = None if sin_exist else num(entrada)
    ps = None if sin_exist else num(salida)
    c = num(cont)
    pu = round(ps / c, 4) if ps and c else None
    # FANNY: si el vendedor no consigue la entrada, se guarda la de lista de la zona (y Precio_Entrada_Lista
    # queda como respaldo para saber de dónde salió). Sin existencia = sin precios.
    pe_final = pe if pe else (None if sin_exist else lista)
    m_abs = round(ps - pe_final, 2) if ps and pe_final else None
    m_pct = round((ps - pe_final) / ps * 100, 1) if ps and pe_final else None   # margen sobre precio de venta
    capturas.append(dict(SKU_ID=sku_id, Tipo=tipo, Producto_ID=prod_id, Comp_Orden=orden,
                         Precio_Entrada=pe_final, Precio_Salida=ps, Margen_Abs=m_abs, Margen_Pct=m_pct,
                         Contenido=cont, UM=um, Precio_UM=pu,
                         Disponible="No" if sin_exist else "Sí", Promo="Sí" if promo else "No",
                         Sell_Out="", Precio_Entrada_Lista=lista, _texto=salida, _texto_e=entrada, _sin=sin_exist, _pe_obs=pe))
    if m_pct is not None:
        st.caption(f"Margen PDV: C$ {m_abs:,.2f} · {m_pct:.1f}%" + ("" if pe else " (con precio de lista)"))
    return pu, promo


def bloque_sku(s):
    comps = m_comp[m_comp["SKU_ID"] == s["SKU_ID"]]
    with st.container(border=True):
        st.markdown(f"#### {s['Producto']}")
        st.caption(f"{s.get('Categoria', '')} · {s.get('Presentacion', '')} · {s['Contenido']} {s['UM']}")
        lista = precio_lista(s)
        pu_own, _ = fila_producto(s["SKU_ID"], "Propio", s["SKU_ID"], s["Producto"],
                                  s["Contenido"], s["UM"], 0,
                                  f" · precio lista {zona_ruta.replace('_', '/').title()} C$ {lista:,.2f}" if lista else "",
                                  lista=lista)
        pu_comp = []
        for _, c in comps.iterrows():
            pu, promo = fila_producto(s["SKU_ID"], "Competencia", c["Comp_ID"],
                                      f"{c['Marca']} · {c['Producto']}", c["Contenido"],
                                      c["UM"], int(c["Comp_Orden"]))
            if pu and not promo:          # el índice excluye precios en promoción
                pu_comp.append(pu)

        if pu_own and pu_comp:
            idx = pu_own / (sum(pu_comp) / len(pu_comp)) * 100
            estado = ("🔴 más caro" if idx > 105 else "🟢 más barato" if idx < 95 else "🟡 paridad")
            st.markdown(f"**Índice de precio: {idx:.0f}** · {estado}")
        else:
            st.caption("Índice: ingrese precio propio y al menos un competidor sin promoción.")
        nota = st.text_input("Dinámica de sell out (opcional)", key=f"so_{fid}_{s['SKU_ID']}",
                             placeholder="Ej.: rota bien, competidor con impulsadora, quiebre frecuente")
        for c in capturas:
            if c["SKU_ID"] == s["SKU_ID"] and c["Tipo"] == "Propio":
                c["Sell_Out"] = nota.strip()


# Categorías: el vendedor elige cuáles levanta en este PDV; cada una va en su pestaña
if "Categoria" not in m_sku or (m_sku["Categoria"] == "").all():
    m_sku["Categoria"] = "General"
categorias = list(dict.fromkeys(m_sku["Categoria"]))       # orden del maestro
elegidas = st.multiselect("Categorías a levantar en este PDV", categorias, default=categorias,
                          key=f"cat_{fid}")
if not elegidas:
    st.info("Elija al menos una categoría.")
    st.stop()
skus_sel = m_sku[m_sku["Categoria"].isin(elegidas)]
pestañas = st.tabs([f"{c} ({(skus_sel['Categoria'] == c).sum()})" for c in elegidas])
for pestaña, cat in zip(pestañas, elegidas):
    with pestaña:
        for _, s in skus_sel[skus_sel["Categoria"] == cat].iterrows():
            bloque_sku(s)

# ───────────────────────── 5. Validación y guardado ─────────────────────────
st.divider()

# Producto con texto no numérico
errores = [f"{c['Producto_ID']}: '{c['_texto']}'" for c in capturas
           if c["_texto"] and not c["_sin"] and c["Precio_Salida"] is None]
errores += [f"{c['Producto_ID']}: entrada '{c['_texto_e']}'" for c in capturas
            if c["_texto_e"] and not c["_sin"] and c["_pe_obs"] is None]
errores += [f"{c['Producto_ID']}: tiene entrada pero falta salida" for c in capturas
            if c["_pe_obs"] and c["Precio_Salida"] is None and not c["_sin"]]
errores += [f"{c['Producto_ID']}: entrada mayor que salida" for c in capturas
            if c["_pe_obs"] and c["Precio_Salida"] and c["_pe_obs"] > c["Precio_Salida"]]
# Propio sin precio ni marca de "sin existencia"
faltan = [c["SKU_ID"] for c in capturas
          if c["Tipo"] == "Propio" and c["Precio_Salida"] is None and not c["_sin"]]

duplicado = not previos[(previos["Cliente_Nombre"].astype(str).map(normalizar) == cliente) &
                        (previos["Fecha"].astype(str) == fecha.isoformat())].empty
confirmar_dup = True
if duplicado:
    st.warning("Este PDV ya tiene un levantamiento con esta fecha.")
    confirmar_dup = st.checkbox("Confirmo que es un levantamiento nuevo", key=f"dup_{fid}")

if st.button("💾 Guardar levantamiento", type="primary", use_container_width=True):
    if errores:
        st.error("Precios no válidos: " + ", ".join(errores))
    elif faltan:
        st.error("Falta precio de salida propio o marcar 'Sin existencia' en: " + ", ".join(faltan))
    elif not confirmar_dup:
        st.error("Confirme el levantamiento duplicado antes de guardar.")
    else:
        ahora = datetime.now(timezone(timedelta(hours=TZ_OFFSET_H)))
        fecha = ahora.date()                       # la fecha guardada es siempre la del momento de guardar
        id_lev = "LV-" + uuid.uuid4().hex[:8].upper()
        filas = []
        for c in capturas:
            if c["Precio_Salida"] is None and not c["_sin"]:
                continue                     # competidor no observado: no se registra
            filas.append({**{k: ("" if v is None else v) for k, v in c.items() if not k.startswith("_")},
                          "ID_Lev": id_lev, "Fecha": fecha.isoformat(),
                          "Hora_Registro": ahora.strftime("%Y-%m-%d %H:%M"),
                          "Usuario": vend.get("Usuario", vend_nombre),
                          "Cliente_ID": pdv["Cliente_ID"], "Cliente_Nombre": pdv["Nombre"],
                          "Canal": pdv["Canal"],
                          "Ruta": pdv["Ruta"], "Municipio": pdv["Municipio"],
                          "Departamento": pdv["Departamento"]})
        try:
            guardar_filas(filas)
            st.session_state.setdefault("guardados", []).append((fecha.isoformat(), pdv["Nombre"]))
            claves_registro.clear()
            st.session_state.form_id += 1
            st.session_state.ok_msg = f"Guardado {id_lev}: {len(filas)} registros de {pdv['Nombre']}."
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo guardar. Revise la conexión e intente de nuevo. Detalle: {e}")
