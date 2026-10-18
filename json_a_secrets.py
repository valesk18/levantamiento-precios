"""Crea .streamlit/secrets.toml a partir de la llave .json de Google Cloud.
Uso: python json_a_secrets.py "FRASE_SECRETA"
"""
import json
import sys
from pathlib import Path

SHEET_ID = "10a3yAmWaHpF3PPCC-dNMerLZEflAvh76V4Use99W_wI"
carpeta = Path.home() / "Documents" / "credenciales"
llaves = sorted(carpeta.glob("levantamiento-precios-*.json"))
if not llaves:
    sys.exit(f"No encontré la llave .json en {carpeta}")
if len(sys.argv) < 2 or len(sys.argv[1]) < 12:
    sys.exit('Indique una frase secreta de al menos 12 caracteres. Ej.: python json_a_secrets.py "frase-larga-2026"')

cred = json.loads(llaves[-1].read_text(encoding="utf-8"))
lineas = [f"SHEET_ID = {json.dumps(SHEET_ID)}", f"PIN_SALT = {json.dumps(sys.argv[1])}", "",
          "[gcp_service_account]"] + [f"{k} = {json.dumps(v)}" for k, v in cred.items()]
destino = Path(__file__).parent / ".streamlit" / "secrets.toml"
destino.write_text("\n".join(lineas) + "\n", encoding="utf-8")
print(f"OK · secrets.toml creado con la llave {llaves[-1].name}")
print(f"OK · Correo de la app: {cred.get('client_email')}")
