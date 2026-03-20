import os
import requests
import subprocess
import sys

SHAREPOINT_URL = os.environ.get("SHAREPOINT_URL")
EXCEL_FILENAME = "DATA_Situations_Mouvement.xlsx"

print("Téléchargement depuis SharePoint...")

# Construire l'URL de téléchargement direct depuis l'URL de partage SharePoint
# Format : remplacer /r/ par le endpoint download
download_url = SHAREPOINT_URL.replace(
    "/:x:/r/", "/:x:/r/"
).split("?")[0]

# L'URL de download direct SharePoint pour un fichier partagé
download_url = f"{download_url.rstrip('/')}?download=1"

print(f"URL download : {download_url}")

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
})

response = session.get(download_url, allow_redirects=True, timeout=60)

print(f"HTTP status : {response.status_code}")
print(f"Content-Type : {response.headers.get('Content-Type', 'inconnu')}")
print(f"Taille : {len(response.content) / 1024:.0f} Ko")

# Vérifier que c'est bien un fichier Excel et pas du HTML
content_type = response.headers.get('Content-Type', '')
if 'html' in content_type.lower():
    print("ERREUR : SharePoint a renvoyé une page HTML — authentification requise")
    print(response.text[:500])
    sys.exit(1)

if response.status_code != 200:
    print(f"ERREUR téléchargement : HTTP {response.status_code}")
    sys.exit(1)

with open(EXCEL_FILENAME, "wb") as f:
    f.write(response.content)

print(f"✓ Fichier téléchargé ({len(response.content) / 1024:.0f} Ko)")

# Lancer transform.py
print("Exécution de transform.py...")
result = subprocess.run(
    ["python3", "transform.py"],
    capture_output=True, text=True
)

if result.returncode != 0:
    print(f"ERREUR transform.py :\n{result.stderr}")
    sys.exit(1)

print("✓ data.json mis à jour avec succès")
