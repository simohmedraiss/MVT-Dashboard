#!/bin/bash
# ================================================
#  Le Mouvement — Mise à jour dashboard en 1 clic
#  Commande : mvtdash
# ================================================

cd "$(dirname "$0")"

PYTHON=/usr/bin/python3

echo "============================================"
echo " Le Mouvement — Mise à jour du dashboard"
echo "============================================"
echo ""

# Trouver le fichier Excel
EXCEL=$(ls *.xlsx 2>/dev/null | head -1)

if [ -z "$EXCEL" ]; then
    echo "❌ Aucun fichier Excel trouvé dans ce dossier."
    echo "   Place ton fichier .xlsx ici et relance."
    read -p "Appuie sur Entrée pour fermer..."
    exit 1
fi

echo "📂 Fichier Excel détecté : $EXCEL"
echo ""

# Générer data.json
echo "⚙️  Génération des données..."
$PYTHON transform.py "$EXCEL"

if [ $? -ne 0 ]; then
    echo "❌ Erreur lors de la génération. Vérifie le fichier Excel."
    read -p "Appuie sur Entrée pour fermer..."
    exit 1
fi

echo ""
echo "📤 Envoi sur GitHub..."

git pull origin main --rebase 2>/dev/null || git pull origin master --rebase 2>/dev/null

# Ajouter tous les fichiers du dashboard
git add data.json Le_Mouvement_Dashboard.html transform.py

git diff --staged --quiet && echo "   Aucun changement détecté." || {
    DATE=$(date '+%Y-%m-%d %H:%M')
    git commit -m "Update dashboard - $DATE"
    git push origin main 2>/dev/null || git push origin master 2>/dev/null
    echo ""
    echo "✅ Dashboard mis à jour en ligne !"
    echo "   https://simohmedraiss.github.io/MVT-Dashboard/Le_Mouvement_Dashboard.html"
}

echo ""
read -p "Appuie sur Entrée pour fermer..."
