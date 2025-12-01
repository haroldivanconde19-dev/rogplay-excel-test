import os
import logging
import requests
import time
from dotenv import load_dotenv
import msal

# Configuración de Logging para ver el output
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Cargar variables
load_dotenv()

# =====================================================
# CONFIGURACIÓN DE MICROSOFT GRAPH (Se asume configurado)
# =====================================================
TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
USER_ID = os.getenv("MS_USER_ID")
FILE_ID = os.getenv("NETFLIX_FILE_ID")
SHEET_NAME = "VENTAS" # Hoja de cálculo a usar para la prueba

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"] 
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Cache de token simplificado
_token_cache = {"access_token": None, "expires_at": 0}

# =====================================================
# FUNCIÓN 1: OBTENER TOKEN (De tu código original)
# =====================================================

def get_token():
    global _token_cache
    
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        logger.error("❌ ERROR: Faltan credenciales MS en .env")
        return None

    # 1. Verificar si ya tenemos un token válido (con 60 seg de margen)
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
# FUNCIÓN 2: ESCRITURA MÍNIMA (PATCH)
# =====================================================

def write_single_cell(file_id: str, sheet_name: str, range_address: str, value: str):
    """
    Intenta escribir una cadena en una celda específica (ej: B1:B1) usando PATCH.
    """
    token = get_token()
    if not token or not file_id:
        return False

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    # Sintaxis más segura para la URL
    url = (
        f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/items/{file_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{range_address}')/values"
    )
    
    # El payload es una lista de listas: [[valor]]
    payload = {"values": [[value]]} 
    
    logger.info(f"💾 Intentando escribir '{value}' en rango: {range_address}")
    
    try:
        response = requests.patch(url, headers=headers, json=payload, timeout=10)
        
        if response.status_code in (200, 202, 204):
            logger.info(f"🎉 ÉXITO: Celda {range_address} actualizada con '{value}'.")
            return True
        else:
            # Imprimimos la respuesta completa del error (esto es crucial)
            logger.error(f"❌ ERROR CRÍTICO ({response.status_code}) al actualizar la celda.")
            logger.error(f"   Response de MS Graph: {response.text}")
            return False
            
    except Exception as e:
        logger.error(f"❌ Excepción fatal al hacer PATCH: {e}")
        return False


# =====================================================
# EJECUCIÓN PRINCIPAL
# =====================================================

if __name__ == "__main__":
    
    # ⚠️ CAMBIO CRÍTICO: Escribir en Z1 para no interferir con los encabezados
    if write_single_cell(FILE_ID, SHEET_NAME, "Z1:Z1", "API_OK"):
        print("\n✅ PRUEBA DE ESCRITURA FINALIZADA CON ÉXITO.")
    else:
        print("\n❌ PRUEBA DE ESCRITURA FALLIDA. Revisa los logs de ERROR para el código 400.")
