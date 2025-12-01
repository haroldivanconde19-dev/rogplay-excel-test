import requests

# ============
# 🔐 CREDENCIALES DE TU APP EN AZURE
# ============
TENANT_ID = "TU_TENANT_ID"
CLIENT_ID = "TU_CLIENT_ID"
CLIENT_SECRET = "TU_CLIENT_SECRET"

# ============
# 📄 DATOS DEL EXCEL EN SHAREPOINT
# ============
SHAREPOINT_SITE = "haroldconde.sharepoint.com"
SITE_NAME = "ROGPLAY"
EXCEL_FILENAME = "test_api_excel.xlsx"
SHEET_NAME = "VENTAS"
READ_RANGE = "A1:A1"
WRITE_RANGE = "B1:B1"
WRITE_VALUE = [["Hola desde Python!"]]

# ============
# 🔐 Obtener access_token
# ============
token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

token_data = {
    "grant_type": "client_credentials",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default"
}

token_resp = requests.post(token_url, data=token_data)
access_token = token_resp.json().get("access_token")

if not access_token:
    print("❌ ERROR: No se pudo obtener el token.")
    exit()

headers = {"Authorization": f"Bearer {access_token}"}

# ============
# 🔎 Obtener SITE-ID de SharePoint
# ============
site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}:/sites/{SITE_NAME}"
site_resp = requests.get(site_url, headers=headers)
site_data = site_resp.json()
site_id = site_data.get("id")

if not site_id:
    print("❌ ERROR: No se encontró el site. Revisa el nombre.")
    print(site_data)
    exit()

print(f"✅ Site ID obtenido: {site_id}")

# ============
# 📄 Obtener ID del archivo Excel
# ============
file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{EXCEL_FILENAME}"
file_resp = requests.get(file_url, headers=headers)
file_data = file_resp.json()
file_id = file_data.get("id")

if not file_id:
    print("❌ ERROR: No se encontró el archivo Excel.")
    print(file_data)
    exit()

print(f"✅ Archivo Excel encontrado. ID: {file_id}")

# ============
# 🔍 Leer celda (READ_RANGE)
# ============
read_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{READ_RANGE}')/values"
read_resp = requests.get(read_url, headers=headers)

if read_resp.status_code == 200:
    read_value = read_resp.json().get("value")
    print(f"📖 Valor en {READ_RANGE}: {read_value}")
else:
    print(f"❌ ERROR al leer {READ_RANGE}")
    print(read_resp.status_code, read_resp.text)

# ============
# ✍️ Escribir valor en celda (WRITE_RANGE)
# ============
write_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{WRITE_RANGE}')"
write_headers = {**headers, "Content-Type": "application/json"}

write_resp = requests.patch(write_url, headers=write_headers, json={"values": WRITE_VALUE})

if write_resp.status_code in [200, 201]:
    print(f"✅ Escrito exitosamente en {WRITE_RANGE}: {WRITE_VALUE}")
else:
    print(f"❌ ERROR al escribir en {WRITE_RANGE}")
    print(write_resp.status_code, write_resp.text)
