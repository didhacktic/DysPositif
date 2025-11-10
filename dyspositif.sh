#!/bin/bash
# -------------------------------------------------
# Dys’Positif – Lancement intelligent (Ubuntu)
# -------------------------------------------------

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$DIR/venv"
ACTIVATE="$VENV_DIR/bin/activate"

echo "=== Dys’Positif – Vérification environnement ==="

# --- Création venv si absent ---
if [ ! -d "$VENV_DIR" ]; then
    echo "[1/5] venv absent → création en cours..."
    python3 -m venv "$VENV_DIR" || { echo "ERREUR : impossible de créer le venv"; exit 1; }
else
    echo "[1/5] venv présent"
fi

# --- Activation ---
source "$ACTIVATE" || { echo "ERREUR : impossible d’activer le venv"; exit 1; }

# --- Mise à jour pip ---
echo "[2/5] Mise à jour pip..."
pip install --upgrade pip

# --- Installation pylirecouleur local ---
echo "[3/5] Installation pylirecouleur (local)..."
if [ -f "$DIR/pylirecouleur/dist/pylirecouleur-0.0.5-py3-none-any.whl" ]; then
    pip install --force-reinstall --no-index \
        "$DIR/pylirecouleur/dist/pylirecouleur-0.0.5-py3-none-any.whl"
    echo "    → pylirecouleur installé"
else
    echo "    → Fichier wheel manquant !"
fi

# --- Installation dépendances PyPI (seulement si absentes) ---
echo "[4/5] Installation dépendances PyPI..."
pip install python-docx==1.1.2 lxml==5.3.0 Pillow==10.4.0 pdfservices-sdk
if [ $? -eq 0 ]; then
    echo "    → Dépendances installées ou déjà présentes"
else
    echo "    → ÉCHEC installation"
    exit 1
fi

# --- PYTHONPATH ---
if ! grep -q "pylirecouleur/src" "$ACTIVATE" 2>/dev/null; then
    echo "[5/5] Ajout PYTHONPATH..."
    echo "export PYTHONPATH=\"$DIR/pylirecouleur/src:\$PYTHONPATH\"" >> "$ACTIVATE"
    echo "    → PYTHONPATH configuré"
else
    echo "[5/5] PYTHONPATH déjà configuré"
fi

# --- Lancement ---
echo ""
echo "=== Lancement de Dys’Positif ==="
python "$DIR/main.py"