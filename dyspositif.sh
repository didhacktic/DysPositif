#!/usr/bin/env bash
# dyspositif – Lancement simplifié Dys’Positif
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$DIR/venv"

# Création du venv + dépendances si absent
if [ ! -d "$VENV" ]; then
    echo "Création de l'environnement virtuel..."
    python3 -m venv "$VENV"
    source "$VENV/bin/activate"
    echo "Mise à jour de pip..."
    pip install --upgrade pip
    echo "Installation des dépendances..."
    pip install python-docx lxml Pillow pdfservices-sdk pylirecouleur
else
    source "$VENV/bin/activate"
fi

# PATCH : Copie le code source embarqué dans site-packages
SITE_PACKAGES="$(python -c 'import site; print(site.getsitepackages()[0])')"
EMBEDDED_SRC="$DIR/pylirecouleur_src/lirecouleur"

if [ -d "$EMBEDDED_SRC" ] && [ ! -d "$SITE_PACKAGES/lirecouleur" ]; then
    echo "Installation locale de pylirecouleur..."
    cp -r "$EMBEDDED_SRC" "$SITE_PACKAGES/"
    echo "pylirecouleur installé localement"
fi

# Lancement
exec python "$DIR/main.py" "$@"