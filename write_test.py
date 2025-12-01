import os
import logging
import requests
import time
from dotenv import load_dotenv
import msal

# ⚙️ Configuración de logs
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 🔄 Cargar variables de entorno
load_dotenv()

# =====================================================
# 🔐 CONFIGURACIÓN
# =====================================================
TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
USER_ID = os.getenv("MS_USER_ID")  # Debe ser el correo (UPN), no el GUID
FILE_ID = os.getenv("NETFLIX_FILE_ID")
SHEET_NAME = "VENTAS"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Cache de token
_token_cache = {"access_token": None, "expires_at": 0}

# =====================================================
# 🔑 TOKEN
# =====================================================

def get_token():
    global _token_cache

    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        logger.error("❌ Faltan credenciales de MS Graph.")
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

# =====================================================
# 📖 LECTURA
# =====================================================

def read_single_cell(file_id: str, sheet_name: str, range_address: str):
    token = get_token()
    if not token or not file_id:
        return None

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    url = (
        f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/items/{file_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{range_address}')/values"
    )

    logger.info(f"💾 Intentando leer rango: {range_address}")

    try:
        response = requests.get(url, headers=headers, timeout=10)

        if response.status_code == 200:
            data = response.json().get("values", [[None]])
            value = data[0][0]
            logger.info(f"🎉 ÉXITO de LECTURA: Celda {range_address} contiene: '{value}'")
            return value
        else:
            logger.error(f"❌ FALLO DE LECTURA ({response.status_code}).")
            logger.error(f"   MS Graph dice: {response.text}")
            return None
    except Exception as e:
        logger.error(f"❌ Excepción al leer: {e}")
        return None

# =====================================================
# ✏️ ESCRITURA
# =====================================================

def write_single_cell(file_id: str, sheet_name: str, range_address: str, value: str):
    token = get_token()
    if not token or not file_id:
        return False

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = (
        f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/items/{file_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{range_address}')/values"
    )

    payload = {"values": [[value]]}

    logger.info(f"📝 Intentando escribir '{value}' en: {range_address}")

    try:
        response = requests.patch(url, headers=headers, json=payload, timeout=10)

        if response.status_code in (200, 202, 204):
            logger.info(f"✅ ESCRITURA OK en {range_address}")
            return True
        else:
            logger.error(f"❌ ERROR DE ESCRITURA ({response.status_code})")
            logger.error(f"   MS Graph dice: {response.text}")
            return False
    except Exception as e:
        logger.error(f"❌ Excepción escribiendo: {e}")
        return False

# =====================================================
# 🚀 PRUEBA
# =====================================================

if __name__ == "__main__":
    print("\n==================================================")
    print("  INICIANDO PRUEBA DOBLE (LECTURA Y ESCRITURA)")
    print("==================================================")

    RANGE_LEER = "A1:A1"
    RANGE_ESCRIBIR = "Z1:Z1"

    read_value = read_single_cell(FILE_ID, SHEET_NAME, RANGE_LEER)

    if read_value is not None:
        success = write_single_cell(FILE_ID, SHEET_NAME, RANGE_ESCRIBIR, "API_TEST_OK")
    else:
        success = False

    if success:
        print("\n✅ PRUEBA COMPLETA: LECTURA Y ESCRITURA OK.")
    elif read_value is not None:
        print("\n⚠️ LECTURA OK, PERO ESCRITURA FALLIDA.")
    else:
        print("\n❌ FALLO DE LECTURA Y ESCRITURA.")
