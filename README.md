# Dys’Positif

Outil open-source qui convertit PDF/DOCX/ODT en documents adaptés dyslexie :

- Coloration syllabique (LireCouleur)
- Coloration des nombres
- Lettres muettes grisées (e, s, ent, ez…)
- Police adaptée, interlignage 1,5, espacement 2,4 pt
- Sortie A3/A4 avec agrandissement tableaux +40 %

## Installation rapide (Ubuntu)

```bash
git clone git@github.com:didhacktic/DysPositif.git
cd DysPositif

python3 -m venv venv
source venv/bin/activate

pip install python-docx adobe-pdfservices-sdk lxml pylirecouleurgit push --force