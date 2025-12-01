import os
import logging
import requests
import time
from dotenv import load_dotenv
import msal

# ⚠️ Nivel DEBUG para ver la URL completa antes de la llamada.
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Cargar variables
load_dotenv()

# =====================================================
# CONFIGURACIÓN DE MICROSOFT GRAPH
# =====================================================
TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
# MS_USER_ID debe ser el ID de Objeto (GUID) del usuario
USER_ID = os.getenv("MS_USER_ID")
FILE_ID = os.getenv("NETFLIX_FILE_ID")
SHEET_NAME = "VENTAS" # Hoja de cálculo confirmada

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"] 
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Cache de token simplificado
_token_cache = {"access_token": None, "expires_at": 0}

# =====================================================
# FUNCIÓN 1: OBTENER TOKEN
# =====================================================

def get_token():
    global _token_cache
    
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        logger.error("❌ ERROR: Faltan credenciales MS en .env/Railway")
        return None

    # 1. Verificar si ya tenemos un token válido
    if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    # 2. Si no hay token o expiró, pedimos uno nuevo
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
            logger.error(f"❌ Error al obtener el token: {result.get('error_description')}")
            return None
    except Exception as e:
        logger.error(f"❌ Excepción obteniendo token: {e}")
        return None

# =====================================================
# FUNCIÓN 2: LECTURA (GET)
# =====================================================

def read_single_cell(file_id: str, sheet_name: str, range_address: str):
    """
    Intenta leer el valor de una celda específica usando GET.
    """
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
            logger.info(f"🎉 ÉXITO de LECTURA: Celda {range_address} contiene el valor: '{value}'")
            return value
        else:
            logger.error(f"❌ FALLO DE LECTURA ({response.status_code}).")
            logger.error(f"   Response de MS Graph: {response.text}")
            return None
            
    except Exception as e:
        logger.error(f"❌ Excepción fatal al hacer GET: {e}")
        return None

# =====================================================
# FUNCIÓN 3: ESCRITURA (PATCH)
# =====================================================

def write_single_cell(file_id: str, sheet_name: str, range_address: str, value: str):
    """
    Intenta escribir una cadena en una celda específica (ej: Z1:Z1) usando PATCH.
    """
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
    
    logger.info(f"💾 Intentando escribir '{value}' en rango: {range_address}")
    
    try:
        response = requests.patch(url, headers=headers, json=payload, timeout=10)
        
        if response.status_code in (200, 202, 204):
            logger.info(f"🎉 ÉXITO de ESCRITURA: Celda {range_address} actualizada.")
            return True
        else:
            logger.error(f"❌ FALLO DE ESCRITURA ({response.status_code}).")
            logger.error(f"   Response de MS Graph: {response.text}")
            return False
            
    except Exception as e:
        logger.error(f"❌ Excepción fatal al hacer PATCH: {e}")
        return False


# =====================================================
# EJECUCIÓN PRINCIPAL
# =====================================================

if __name__ == "__main__":
    
    print("\n==================================================")
    print("      INICIANDO PRUEBA DOBLE (LECTURA Y ESCRITURA)")
    print("==================================================")
    
    # 1. Prueba de LECTURA (GET a A6)
    read_value = read_single_cell(FILE_ID, SHEET_NAME, "A6:A6")

    # 2. Prueba de ESCRITURA (PATCH a Z1)
    if read_value is not None:
        write_success = write_single_cell(FILE_ID, SHEET_NAME, "Z1:Z1", "API_OK_FINAL")
    else:
        write_success = False

    if write_success:
        print("\n✅ PRUEBA COMPLETA: Ambos tests fueron exitosos.")
    elif read_value is not None and not write_success:
        print("\n⚠️ RESULTADO AMBIGUO: LECTURA OK, ESCRITURA FALLIDA.")
        print("   CAUSA PROBABLE: Error de sintaxis o tipo de dato, o permiso de escritura restringido.")
    else:
        print("\n❌ PRUEBA FALLIDA: FALLO DE LECTURA Y ESCRITURA.")
        print("   CAUSA PROBABLE: MS_USER_ID, NETFLIX_FILE_ID o SHEET_NAME es incorrecto.")
