#!/usr/bin/env bash
# dyspositif – Lancement simplifié Dys’Positif
# Ce script crée/active un venv et installe les dépendances manquantes de la même façon
# pour tous les paquets (y compris pylirecouleur -> importable via 'lirecouleur').
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$DIR/venv"

# Helper : test si un module est importable dans l'environnement python courant
_module_installed() {
    module="$1"
    python -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$module') else 1)"
}

# -------------------------------------------------------------------------
# 1) Création / activation du venv
# -------------------------------------------------------------------------
NEW_VENV=0
if [ ! -d "$VENV" ]; then
    echo "Création de l'environnement virtuel..."
    python3 -m venv "$VENV"
    NEW_VENV=1
fi

# Activer le venv (assure que 'python' ci‑dessous pointe vers le venv)
# shellcheck source=/dev/null
source "$VENV/bin/activate"

# Mettre pip/setuptools/wheel à jour
echo "Mise à jour de pip, setuptools et wheel..."
python -m pip install --upgrade pip setuptools wheel

# -------------------------------------------------------------------------
# 2) Dépendances "principales" : n'installer que les paquets/modules manquants
#    pylirecouleur est traité comme les autres : distribution 'pylirecouleur',
#    module importable 'lirecouleur'.
# -------------------------------------------------------------------------
declare -A DEPS
DEPS=(
  ["docx"]="python-docx"
  ["lxml"]="lxml"
  ["PIL"]="Pillow"
  ["adobe"]="pdfservices-sdk"
  ["spacy"]="spacy"
  ["lirecouleur"]="pylirecouleur"
)

check_pkg() {
  mod="$1"
  dist="$2"
  python - <<PY
import importlib, importlib.metadata, sys
try:
    importlib.import_module("$mod")
    sys.exit(0)
except Exception:
    pass
try:
    importlib.metadata.version("$dist")
    sys.exit(0)
except Exception:
    sys.exit(1)
PY
}

MISSING_PKGS=()
for mod in "${!DEPS[@]}"; do
    if check_pkg "$mod" "${DEPS[$mod]}"; then
        echo "Module/distribution '${mod}'/'${DEPS[$mod]}' présent."
    else
        echo "Module '${mod}' absent → marquage pour installation (${DEPS[$mod]})."
        MISSING_PKGS+=("${DEPS[$mod]}")
    fi
done

if [ ${#MISSING_PKGS[@]} -gt 0 ]; then
    echo "Installation des paquets manquants : ${MISSING_PKGS[*]} ..."
    python -m pip install --upgrade "${MISSING_PKGS[@]}"
else
    echo "Toutes les dépendances principales sont déjà installées."
fi

# Si on vient juste de créer le venv, télécharger le modèle spaCy (unique fois)
if [ "$NEW_VENV" -eq 1 ]; then
    echo "Téléchargement du modèle spaCy fr_core_news_md (uniquement après création du venv)..."
    if python -m spacy download fr_core_news_md; then
        echo "Modèle spaCy fr_core_news_md installé."
    else
        echo "Attention : échec du téléchargement du modèle spaCy fr_core_news_md."
        echo "Vous pouvez l'installer manuellement : python -m spacy download fr_core_news_md"
    fi
fi

# -------------------------------------------------------------------------
# 3) Pas de vérification finale spécifique pour pylirecouleur : tout est traité
#    de la même manière que les autres paquets (voir DEPS).
# -------------------------------------------------------------------------

# -------------------------------------------------------------------------
# 4) Lancement
# -------------------------------------------------------------------------
exec python "$DIR/main.py" "$@"