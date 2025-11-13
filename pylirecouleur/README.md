### pylirecouleur — Build depuis les sources

Ce document explique comment construire la wheel (fichier .whl) de pylirecouleur à partir des sources présentes dans ce répertoire.

Pré‑requis
- Python 3.8+ installé (ou la version compatible que vous utilisez pour DysPositif).
- pip disponible.
- (Optionnel) un virtualenv pour isoler l'environnement de build.

Étapes rapides (recommandé depuis le répertoire pylirecouleur)
# depuis la racine du dépôt :
cd pylirecouleur

# mettre à jour l'outil de build
python -m pip install --upgrade build

# nettoyer un ancien dist/ (optionnel)
rm -rf dist/

# construire la wheel et la placer dans pylirecouleur/dist/
python -m build --wheel --outdir dist .


La commande ci‑dessus crée un fichier du type :
dist/pylirecouleur-<version>+<tag>-py3-none-any.whl


Installer et tester localement
# installer la wheel construite dans l'environnement actif
python -m pip install --upgrade dist/*.whl

# vérification rapide dans le Python du venv
python -c "import importlib.util, sys; print('ok' if importlib.util.find_spec('lirecouleur') else 'missing')"

## Remarques utiles
- Le package exposé pour import est `lirecouleur` (c'est le nom du module utilisé par DysPositif).
- Pour que `python -m build` fonctionne, ce répertoire doit contenir les métadonnées de packaging : un `pyproject.toml` (préféré) ou `setup.cfg` / `setup.py`. 
Exemple :
```toml
[build-system]
requires = ["setuptools>=61", "wheel"]
build-backend = "setuptools.build_meta"
```
Placez‑le à la racine de `pylirecouleur/` avant de lancer le build.

# Licence
- Le code source hérité de ce répertoire est soumis à la licence fournie à la racine du dépôt (voir ../LICENSE).
