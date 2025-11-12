#!/usr/bin/env bash
# dyspositif – Lancement simplifié Dys’Positif
# Ce script crée/active un venv, installe les dépendances manquantes proprement
# et installe le package vendored `pylirecouleur` en priorité via une wheel locale
# (pylirecouleur/dist/*.whl). Si aucune wheel locale n'est trouvée, il installe la
# source embarquée ou retombe sur PyPI.
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
#    Mapping : <importable module name> => <pip package name>
#    Remarque : on vérifie d'abord l'import du module, puis la présence de la
#    distribution (importlib.metadata.version) pour éviter les confusions
#    entre le nom d'import et le nom de distribution PyPI.
# -------------------------------------------------------------------------

declare -A DEPS
DEPS=(
  ["docx"]="python-docx"
  ["lxml"]="lxml"
  ["PIL"]="Pillow"
  ["adobe"]="pdfservices-sdk"   # top-level package fourni par pdfservices-sdk
  ["spacy"]="spacy"
)

# check_pkg modulaire: retourne 0 si présent (import ok OU distribution installée), 1 sinon.
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
    # tentative silencieuse mais informative
    if python -m spacy download fr_core_news_md; then
        echo "Modèle spaCy fr_core_news_md installé."
    else
        echo "Attention : échec du téléchargement du modèle spaCy fr_core_news_md."
        echo "Vous pouvez l'installer manuellement : python -m spacy download fr_core_news_md"
    fi
fi

# -------------------------------------------------------------------------
# 3) Gestion prioritaire via wheel pour pylirecouleur (package 'lirecouleur')
# -------------------------------------------------------------------------
EMBEDDED_DIR="$DIR/pylirecouleur"
WHEEL_GLOB="$EMBEDDED_DIR/dist/*.whl"
PKG_NAME="lirecouleur"

echo "Vérification du package '${PKG_NAME}'..."
if _module_installed "${PKG_NAME}"; then
    echo "'${PKG_NAME}' déjà disponible dans l'environnement."
else
    # 3.a) Priorité : installer une wheel locale si elle existe
    if compgen -G "$WHEEL_GLOB" > /dev/null; then
        echo "Wheel locale détectée dans ${EMBEDDED_DIR}/dist/ -> installation prioritaire..."
        # utilise --force-reinstall pour garantir la wheel locale
        python -m pip install --upgrade --force-reinstall "$EMBEDDED_DIR/dist/"*.whl
        echo "Installation depuis wheel locale terminée."
    else
        # 3.b) Pas de wheel : tenter d'installer depuis la source embarquée (pylirecouleur/)
        if [ -d "$EMBEDDED_DIR" ]; then
            echo "Source embarquée détectée : $EMBEDDED_DIR"
            if [ "${PYLIRE_DEV:-0}" = "1" ]; then
                echo "Mode développement demandé (PYLIRE_DEV=1) : installation editable..."
                if python -m pip install -e "$EMBEDDED_DIR"; then
                    echo "Installation editable réussie pour pylirecouleur."
                else
                    echo "ERREUR : l'installation editable de pylirecouleur a échoué."
                    echo "Vérifiez la présence d'un setup.cfg / pyproject.toml dans $EMBEDDED_DIR"
                    echo "et exécutez : python -m pip install -e $EMBEDDED_DIR  manuellement pour diagnostic."
                fi
            else
                echo "Installation via pip depuis la source embarquée..."
                if python -m pip install --upgrade "$EMBEDDED_DIR"; then
                    echo "pylirecouleur installé depuis la source embarquée."
                else
                    echo "ERREUR : échec de l'installation de pylirecouleur depuis $EMBEDDED_DIR via pip."
                    echo "Vérifiez que $EMBEDDED_DIR contient pyproject.toml ou setup.cfg + README/LICENSE."
                    echo "Vous pouvez aussi installer en mode dev : export PYLIRE_DEV=1 ; ./dyspositif.sh"
                fi
            fi
        else
            # 3.c) Pas de source embarquée : fallback PyPI
            echo "Aucune source embarquée trouvée. Tentative d'installation depuis PyPI (fallback)..."
            if python -m pip install --upgrade pylirecouleur; then
                echo "Tentative PyPI réussie — vérifiez que le package n'est pas 'vide'."
            else
                echo "ATTENTION : l'installation depuis PyPI a échoué ou le package PyPI est défectueux."
                echo "Si le package PyPI est problématique, récupérez la source valide de pylirecouleur"
                echo "et placez-la dans : $EMBEDDED_DIR  puis relancez ce script."
            fi
        fi
    fi
fi

# Vérification finale : alerte si le package n'est toujours pas disponible
if _module_installed "${PKG_NAME}"; then
    echo "'${PKG_NAME}' prêt."
else
    echo "ERREUR FINALE : le package '${PKG_NAME}' n'est pas disponible dans l'environnement python."
    echo "Consignes :"
    echo " - Si vous avez la source pylirecouleur, placez-la dans : $EMBEDDED_DIR"
    echo " - Assurez-vous que le dossier contient un pyproject.toml et/ou setup.cfg"
    echo " - Pour le développement, activez : export PYLIRE_DEV=1 ; ./dyspositif.sh"
    echo "L'application peut fonctionner sans 'lirecouleur' si vous désactivez l'option 'syllabes' dans l'UI."
fi

# -------------------------------------------------------------------------
# 4) Lancement
# -------------------------------------------------------------------------
exec python "$DIR/main.py" "$@"