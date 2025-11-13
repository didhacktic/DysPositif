#!/usr/bin/env bash
# dyspositif – Lancement simplifié Dys’Positif
# Ce script crée/active un venv, installe les dépendances manquantes proprement
# et installe le package pylirecouleur uniquement depuis l'asset Release GitHub
# si le module n'est pas déjà présent dans le venv.
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
# -------------------------------------------------------------------------
declare -A DEPS
DEPS=(
  ["docx"]="python-docx"
  ["lxml"]="lxml"
  ["PIL"]="Pillow"
  ["adobe"]="pdfservices-sdk"
  ["spacy"]="spacy"
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
# 3) Installation pylirecouleur :
#    - si le module 'lirecouleur' est déjà présent -> ne rien faire
#    - sinon -> installer depuis l'URL de la Release GitHub (wheel prébuild)
# -------------------------------------------------------------------------
PKG_NAME="lirecouleur"
WHEEL_RELEASE_URL="https://github.com/didhacktic/DysPositif/releases/download/v0.0.5-dysp1/pylirecouleur-0.0.5+dysp1-py3-none-any.whl"

echo "Vérification du package '${PKG_NAME}'..."
if _module_installed "${PKG_NAME}"; then
    echo "'${PKG_NAME}' déjà disponible dans l'environnement → aucune action requise."
else
    echo "'${PKG_NAME}' absent → installation depuis la Release GitHub :"
    echo "  $WHEEL_RELEASE_URL"
    if python -m pip install --upgrade "$WHEEL_RELEASE_URL"; then
        echo "Installation de pylirecouleur depuis la Release réussie."
    else
        echo "ERREUR : échec de l'installation de pylirecouleur depuis la Release GitHub."
        echo "Le script n'essaie pas de builder localement ni d'utiliser PyPI."
    fi
fi

# Vérification finale : alerte si le package n'est toujours pas disponible
if _module_installed "${PKG_NAME}"; then
    echo "'${PKG_NAME}' prêt."
else
    echo "ERREUR FINALE : le package '${PKG_NAME}' n'est pas disponible dans l'environnement python."
    echo "Consignes :"
    echo " - Si l'installation depuis la Release a échoué, installez manuellement la wheel :"
    echo "     python -m pip install \"$WHEEL_RELEASE_URL\""
fi

# -------------------------------------------------------------------------
# 4) Lancement
# -------------------------------------------------------------------------
exec python "$DIR/main.py" "$@"