#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 
# Syllabation via pylirecouleur (Arkaline, GPL v3)
# Source originale : https://framagit.org/arkaline/pylirecouleur
# Copié et adapté pour DysPositif (GPL v3)

import re
import os
from docx.shared import RGBColor
from lirecouleur.word import syllables

# Ajout du chemin src du dépôt pylirecouleur
pylire_root = os.path.join(os.path.dirname(__file__), '..', 'pylirecouleur')
if os.path.exists(pylire_root):
    src_path = os.path.join(pylire_root, 'src')
    if src_path not in sys.path:
        sys.path.insert(0, src_path)

# Import direct des fonctions
from lirecouleur.word import syllables  # Fonction principale

COUL_SYLL = [RGBColor(220, 20, 60), RGBColor(30, 144, 255)]
WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ'’'-]+")

# Normalisation pour lirecouleur (supprime accents)
ACCENT_MAP = str.maketrans({
    'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
    'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
    'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
    'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y',
    'æ': 'ae', 'œ': 'oe',
    'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A', 'Å': 'A',
    'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
    'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
    'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O',
    'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U',
    'Ç': 'C', 'Ñ': 'N', 'Ý': 'Y', 'Ÿ': 'Y',
    'Æ': 'AE', 'Œ': 'OE'
})

def normalize(word: str) -> str:
    return word.lower().translate(ACCENT_MAP)

def apply_syllables(doc):
    counter = [0]
    containers = [doc]
    for table in doc.tables:
        containers.extend(table._cells)
    for shape in doc.inline_shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            containers.append(shape.text_frame)

    for container in containers:
        for p in container.paragraphs:
            if not p.text.strip():
                continue
            texte = p.text
            p.clear()
            i = 0
            while i < len(texte):
                c = texte[i]
                if c.isspace() or c in ".,;:!?()[]{}«»“”'’/\\-–—*+=<>@#$%^&~":
                    p.add_run(c)
                    i += 1
                    continue
                match = WORD_PATTERN.match(texte[i:])
                if match:
                    mot = match.group()
                    mot_norm = normalize(mot)

                    # Segmentation via pylirecouleur.syllables()
                    try:
                        syll_parts = syllables(mot_norm)
                    except Exception:
                        syll_parts = [mot_norm]  # Fallback silencieux

                    # Alignement syllabes → graphie originale
                    pos = 0
                    for part in syll_parts:
                        part_len = len(part)
                        orig_part = mot[pos:pos + part_len]
                        # Ajustement si décalage accent
                        if normalize(orig_part) != part:
                            orig_part = mot[pos:pos + 1]
                            pos += 1
                        else:
                            pos += part_len
                        if orig_part:
                            run = p.add_run(orig_part)
                            run.font.color.rgb = COUL_SYLL[counter[0] % 2]
                            counter[0] += 1

                    i += len(mot)
                else:
                    p.add_run(c)
                    i += 1