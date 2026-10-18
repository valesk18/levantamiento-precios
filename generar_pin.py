"""
Genera un PIN aleatorio de 4 dígitos por RUTA (hoja Maestro_Ruta) y crea:
    PINES_Levantamiento_Precios.xlsx  (en esta misma carpeta)
      - Hoja "Entrega_PIN": Ruta, Descripcion, PIN   -> para entregar a cada ruta
      - Hoja "Para_Sheet" : Ruta, PIN_Hash           -> lo que va en Maestro_Ruta

Fuente de rutas:
    - Si existe .streamlit/secrets.toml -> lee Maestro_Ruta del Google Sheet
    - Si no                            -> lee data/Maestro_Ruta.csv (modo local)

Uso (con valesK activo, desde la carpeta de la app):
    python generar_pin.py               -> genera PIN y el Excel
    python generar_pin.py --subir       -> además escribe PIN_Hash en el Google Sheet
                                           (o en data/Maestro_Ruta.csv si está en modo local)
    python generar_pin.py --solo R-05   -> repone el PIN de una sola ruta

⚠ El Excel contiene PIN en texto claro: entréguelos y luego bórrelo o guárdelo fuera del proyecto.
"""
import hashlib
import secrets as rnd
import sys
import tomllib
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE = Path(__file__).parent
RUTA_SECRETS = BASE / ".streamlit" / "secrets.toml"
RUTA_CSV = BASE / "data" / "Maestro_Ruta.csv"
SALIDA = BASE / "PINES_Levantamiento_Precios.xlsx"
PIN_DEBILES = {"0000", "1111", "2222", "3333", "4444", "5555", "6666", "7777",
               "8888", "9999", "1234", "4321", "1212", "0123", "9876"}

cfg = tomllib.loads(RUTA_SECRETS.read_text(encoding="utf-8")) if RUTA_SECRETS.exists() else {}
SAL = cfg.get("PIN_SALT", "")
USAR_SHEET = "gcp_service_account" in cfg and "SHEET_ID" in cfg


def hash_pin(pin: str) -> str:
    return hashlib.sha256((SAL + pin).encode("utf-8")).hexdigest()


def pin_nuevo(usados: set) -> str:
    while True:
        p = f"{rnd.randbelow(10000):04d}"
        if p not in PIN_DEBILES and p not in usados:
            usados.add(p)
            return p


def hoja_rutas():
    import gspread
    gc = gspread.service_account_from_dict(cfg["gcp_service_account"])
    return gc.open_by_key(cfg["SHEET_ID"]).worksheet("Maestro_Ruta")


# 1. Leer rutas
if USAR_SHEET:
    rutas = pd.DataFrame(hoja_rutas().get_all_records()).astype(str)
    origen = "Google Sheet"
else:
    rutas = pd.read_csv(RUTA_CSV, dtype=str).fillna("")
    origen = "data/Maestro_Ruta.csv (modo local)"
if "Descripcion" not in rutas:
    rutas["Descripcion"] = ""

if "--solo" in sys.argv:
    ruta = sys.argv[sys.argv.index("--solo") + 1]
    rutas = rutas[rutas["Ruta"] == ruta]
    if rutas.empty:
        sys.exit(f"No existe la ruta {ruta} en Maestro_Ruta.")

# 2. Generar PIN y hash
usados: set = set()
rutas["PIN"] = [pin_nuevo(usados) for _ in range(len(rutas))]
rutas["PIN_Hash"] = rutas["PIN"].map(hash_pin)

# 3. Excel
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    rutas[["Ruta", "Descripcion", "PIN"]].to_excel(xw, sheet_name="Entrega_PIN", index=False)
    rutas[["Ruta", "PIN_Hash"]].to_excel(xw, sheet_name="Para_Sheet", index=False)
    xw.sheets["Entrega_PIN"].column_dimensions["A"].width = 10
    xw.sheets["Entrega_PIN"].column_dimensions["B"].width = 28
    xw.sheets["Para_Sheet"].column_dimensions["B"].width = 68
    for celda in xw.sheets["Entrega_PIN"]["C"][1:]:
        celda.number_format = "@"  # conserva ceros a la izquierda (ej. 0482)
    xw.sheets["Entrega_PIN"]["E1"] = (f"Generado: {datetime.now():%Y-%m-%d %H:%M} · "
                                      f"Sal: {'sí' if SAL else 'NO (modo local)'}")

# 4. Opcional: escribir PIN_Hash en la fuente
if "--subir" in sys.argv and not USAR_SHEET:
    todos = pd.read_csv(RUTA_CSV, dtype=str).fillna("")
    nuevos = dict(zip(rutas["Ruta"], rutas["PIN_Hash"]))
    todos["PIN_Hash"] = todos["Ruta"].map(nuevos).fillna(todos["PIN_Hash"])
    todos.to_csv(RUTA_CSV, index=False, encoding="utf-8")
    print("PIN_Hash actualizado en data/Maestro_Ruta.csv (modo local).")
elif "--subir" in sys.argv:
    ws = hoja_rutas()
    encabezados = ws.row_values(1)
    if "PIN_Hash" not in encabezados:
        sys.exit("La hoja Maestro_Ruta no tiene la columna PIN_Hash.")
    col = encabezados.index("PIN_Hash") + 1
    lista = ws.col_values(encabezados.index("Ruta") + 1)
    ws.batch_update([{"range": f"{chr(64 + col)}{lista.index(r) + 1}", "values": [[h]]}
                     for r, h in zip(rutas["Ruta"], rutas["PIN_Hash"])])
    print("PIN_Hash actualizado en el Google Sheet.")

print(f"Rutas leídas de: {origen}")
print(f"Sal usada: {'PIN_SALT de secrets.toml' if SAL else 'NINGUNA (modo local, no usar en producción)'}")
print(f"{len(rutas)} PIN generados -> {SALIDA.name}")
