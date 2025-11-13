# Dys’Positif

Outil open-source qui convertit PDF/DOCX/ODT en documents adaptés dyslexie :

- Coloration syllabique (LireCouleur)
- Coloration des nombres
- Lettres muettes grisées (e, s, ent, ez…)
- Police adaptée, interlignage 1,5, espacement 2,4 pt
- Sortie A3/A4 avec agrandissement tableaux +40 %

## Installation

### 1. Cloner le projet
git clone https://github.com/didhacktic/DysPositif.git ~/didhacktic/DysPositif
cd ~/didhacktic/DysPositif

### 2. Rendre le script exécutable
chmod +x dyspositif.sh

### 3. Lancement
./dyspositif.sh

## Paquet tiers embarqué : pylirecouleur (wheel)

Ce dépôt fournit, pour commodité des utilisateurs, une copie binaire (wheel) du paquet tiers
pylirecouleur (LireCouleur) qui n'est plus disponible sur PyPI.

Emplacement du wheel inclus
https://github.com/didhacktic/DysPositif/releases/tag/v0.0.5-dysp1

Le script de lancement de l'application dyspositif.sh installe automatiquent cette version de pylirecouleur.

Important — licence et obligations
- pylirecouleur est un logiciel libre sous la licence GNU General Public License version 3 (ou ultérieure) — GPL-3.0-or-later.
- Le fichier LICENSE à la racine du dépôt contient le texte complet de la licence GPL v3.
- Le fichier NOTICE fournit l'attribution détaillée et indique l'URL du projet upstream :
  https://framagit.org/arkaline/pylirecouleur

Conformité GPL — notes pour les redistributions
- En redistribuant le wheel, nous respectons la GPL v3 en incluant la licence complète (LICENSE)
  et en indiquant la provenance du logiciel (NOTICE). Si vous republiez ce dépôt ou le wheel,
  vous devez respecter la GPL : fournir le texte de la licence, indiquer les modifications et
  rendre disponible le code source (voir la page upstream pour le code source original).
- Si vous modifiez pylirecouleur, les modifications doivent être mises à disposition sous la
  même licence (GPL v3+), et les fichiers modifiés doivent être marqués comme tels (notice de changement).
