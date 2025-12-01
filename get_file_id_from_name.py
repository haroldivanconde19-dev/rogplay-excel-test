import os
import requests
from dotenv import load_dotenv
from get_token import get_token

load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
USER_ID = os.getenv("MS_USER_ID")
EXCEL_FILENAME = "test_api_excel.xlsx"  # Cambia por el nombre exacto de tu archivo

def find_file_id(filename: str):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{GRAPH_BASE_URL}/users/{USER_ID}/drive/root/search(q='{filename}')"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        items = response.json().get("value", [])
        for item in items:
            if item["name"].lower() == filename.lower():
                print(f"\n✅ FILE_ID encontrado: {item['id']}")
                print(f"🧠 FULL NAME: {item['name']}")
                print(f"📁 PATH: {item['parentReference']['path']}")
                return item["id"]
        print("\n❌ Archivo no encontrado.")
    else:
        print(f"\n❌ Error en búsqueda: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    find_file_id(EXCEL_FILENAME)
