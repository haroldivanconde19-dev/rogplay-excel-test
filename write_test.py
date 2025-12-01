import os
import logging
import requests
import time
from dotenv import load_dotenv
import msal

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
load_dotenv()

TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
USER_ID = os.getenv("MS_USER_ID")
SHEET_NAME = "VENTAS"
EXCEL_FILE_NAME = "test_api_excel.xlsx"  # Nombre exacto del archivo

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
_token_cache = {"access_token": None, "expires_at": 0}

def get_token():
    global _token_cache

    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        logger.error("❌ ERROR: Faltan credenciales MS en .env")
        return None

    if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    try:
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )
        logger.info("🔄 Solicitando nuevo token a Microsoft...")
        result = app.acquire_token_for_client(scopes=SCOPE)

        if "access_token" in result:
            logger.info("✅ Token obtenido y guardado en caché.")
            _token_cache["access_token"] = result["access_token"]
            _token_cache["expires_at"] = time.time() + result.get("expires_in", 3599)
            return result["access_token"]
        else:
            logger.error(f"❌ Error al obtener token: {result.get('error_description')}")
            return None
    except Exception as e:
        logger.error(f"❌ Excepción obteniendo token: {e}")
        return None

def read_single_cell(sheet_name: str, range_address: str):
    token = get_token()
    if not token:
        return None

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    url = (
        f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/root:/{EXCEL_FILE_NAME}:/workbook"
        f"/worksheets('{sheet_name}')/range(address='{range_address}')/values"
    )

    logger.info(f"💾 Intentando leer rango: {range_address}")
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json().get("values", [[None]])
            value = data[0][0]
            logger.info(f"🎉 ÉXITO de LECTURA: Celda {range_address} contiene el valor: '{value}'")
            return value
        else:
            logger.error(f"❌ FALLO DE LECTURA ({response.status_code}).")
            logger.error(f"   MS Graph dice: {response.text}")
            return None
    except Exception as e:
        logger.error(f"❌ Excepción fatal al hacer GET: {e}")
        return None

def write_single_cell(sheet_name: str, range_address: str, value: str):
    token = get_token()
    if not token:
        return False

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = (
        f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/root:/{EXCEL_FILE_NAME}:/workbook"
        f"/worksheets('{sheet_name}')/range(address='{range_address}')/values"
    )

    payload = {"values": [[value]]}
    logger.info(f"💾 Intentando escribir '{value}' en rango: {range_address}")

    try:
        response = requests.patch(url, headers=headers, json=payload, timeout=10)
        if response.status_code in (200, 202, 204):
            logger.info(f"🎉 ÉXITO de ESCRITURA: Celda {range_address} actualizada.")
            return True
        else:
            logger.error(f"❌ FALLO DE ESCRITURA ({response.status_code}).")
            logger.error(f"   MS Graph dice: {response.text}")
            return False
    except Exception as e:
        logger.error(f"❌ Excepción fatal al hacer PATCH: {e}")
        return False

if __name__ == "__main__":
    print("\n==================================================")
    print("  INICIANDO PRUEBA DOBLE (LECTURA Y ESCRITURA)")
    print("==================================================")

    read_value = read_single_cell(SHEET_NAME, "A1:A1")
    if read_value is not None:
        write_success = write_single_cell(SHEET_NAME, "Z1:Z1", "API_OK_FINAL")
    else:
        write_success = False

    if write_success:
        print("\n✅ PRUEBA COMPLETA: Ambos tests fueron exitosos.")
    elif read_value is not None and not write_success:
        print("\n⚠️ RESULTADO AMBIGUO: LECTURA OK, ESCRITURA FALLIDA.")
    else:
        print("\n❌ FALLO DE LECTURA Y ESCRITURA.")
