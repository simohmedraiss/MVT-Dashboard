import os
import requests
import subprocess
import sys

SHAREPOINT_URL = os.environ.get("SHAREPOINT_URL")
EXCEL_FILENAME = "DATA_Situations_Mouvement.xlsx"

# Étape 1 : Télécharger le fichier
print("Téléchargement depuis SharePoint...")

download_url = SHAREPOINT_URL.replace(
    "_layouts/15/onedrive.aspx",
    "_layouts/15/download.aspx"
)

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; GitHub Actions)"
})

response = session.get(download_url, allow_redirects=True, timeout=60)

if response.status_code != 200:
    print(f"ERREUR téléchargement : HTTP {response.status_code}")
    sys.exit(1)

with open(EXCEL_FILENAME, "wb") as f:
    f.write(response.content)

print(f"✓ Fichier téléchargé ({len(response.content) / 1024:.0f} Ko)")

# Étape 2 : Lancer transform.py
print("Exécution de transform.py...")
result = subprocess.run(
    ["python3", "transform.py"],
    capture_output=True, text=True
)

if result.returncode != 0:
    print(f"ERREUR transform.py :\n{result.stderr}")
    sys.exit(1)

print("✓ data.json mis à jour avec succès")
