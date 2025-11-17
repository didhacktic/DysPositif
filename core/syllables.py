#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# Syllabation via pylirecouleur (Arkaline, GPL v3)
# Source originale : https://framagit.org/arkaline/pylirecouleur
# Copié et adapté pour DysPositif (GPL v3)
#
# Modifications apportées :
# - Ajout de l'import sys avant manipulation de sys.path (corrige NameError potentiel)
# - Consolidation de l'import de la fonction `syllables` pour éviter la redondance
# - Clarification : WORD_PATTERN ne contient PAS le tiret (-). Le tiret est traité
#   comme séparateur afin d'être cohérent avec le traitement des noms composés
#   (ex. "Jean-Michel") et éviter des recollages indésirables par d'autres modules.
#

import sys
import re
import os
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# Ajout du chemin src du dépôt pylirecouleur (si présent dans l'arborescence du projet)
# On insère ce chemin au début de sys.path afin que l'import suivant trouve le module.
pylire_root = os.path.join(os.path.dirname(__file__), '..', 'pylirecouleur')
if os.path.exists(pylire_root):
    src_path = os.path.join(pylire_root, 'src')
    if src_path not in sys.path:
        sys.path.insert(0, src_path)

# Import direct de la fonction de segmentation syllabique depuis lirecouleur.
# Important : cet import doit intervenir après la modification éventuelle de sys.path.
from lirecouleur.word import syllables  # Fonction principale de segmentation

# Couleurs utilisées pour l'alternance syllabique
COUL_SYLL = [RGBColor(220, 20, 60), RGBColor(30, 144, 255)]

# Pattern acceptant lettres latines, lettres accentuées françaises, apostrophes typographiques, etc.
# NOTE : le tiret '-' a été retiré de la classe afin que les mots liés par un tiret
# soient traités comme mot + séparateur + mot (cohérent avec la logique de parsing char-by-char).
WORD_PATTERN = re.compile(r"[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF'\u2019]+")

# Table de normalisation pour supprimer/mapper les accents avant la segmentation
ACCENT_MAP = str.maketrans({
    '\u00e0': 'a', '\u00e1': 'a', '\u00e2': 'a', '\u00e3': 'a', '\u00e4': 'a', '\u00e5': 'a',
    '\u00e8': 'e', '\u00e9': 'e', '\u00ea': 'e', '\u00eb': 'e',
    '\u00ec': 'i', '\u00ed': 'i', '\u00ee': 'i', '\u00ef': 'i',
    '\u00f2': 'o', '\u00f3': 'o', '\u00f4': 'o', '\u00f5': 'o', '\u00f6': 'o',
    '\u00f9': 'u', '\u00fa': 'u', '\u00fb': 'u', '\u00fc': 'u',
    '\u00e7': 'c', '\u0153': 'oe', '\u00e6': 'ae',
    '\u00c0': 'A', '\u00c1': 'A', '\u00c2': 'A', '\u00c3': 'A', '\u00c4': 'A', '\u00c5': 'A',
    '\u00c8': 'E', '\u00c9': 'E', '\u00ca': 'E', '\u00cb': 'E',
    '\u00cc': 'I', '\u00cd': 'I', '\u00ce': 'I', '\u00cf': 'I',
    '\u00d2': 'O', '\u00d3': 'O', '\u00d4': 'O', '\u00d5': 'O', '\u00d6': 'O',
    '\u00d9': 'U', '\u00da': 'U', '\u00db': 'U', '\u00dc': 'U',
    '\u00c7': 'C', '\u00d1': 'N', '\u00dd': 'Y', '\u0178': 'Y',
    '\u00c6': 'AE', '\u0152': 'OE'
})

def normalize(word: str) -> str:
    """Retourne une version normalisée (minuscule, accents mappés) du mot."""
    return word.lower().translate(ACCENT_MAP)

def _iter_vml_textbox_paragraphs(doc):
    """
    Itère les paragraphes contenus dans les zones de texte VML (legacy Word text boxes).
    
    Les zones de texte VML sont stockées dans la structure XML suivante :
    w:pict → v:shape → v:textbox → w:txbxContent → w:p (paragraphs)
    
    Ces zones ne sont pas accessibles via doc.inline_shapes et nécessitent
    un parcours direct de la structure XML.
    """
    # Namespace URIs pour VML (pas dans le nsmap par défaut de python-docx)
    VML_NS = 'urn:schemas-microsoft-com:vml'
    
    # Rechercher dans tout le document (pas seulement le body)
    # car les zones VML peuvent être n'importe où
    root = doc.element
    
    # Parcourir tous les éléments w:txbxContent (contenu de textbox)
    # C'est plus robuste que de chercher w:pict car cela trouve tous les textboxes VML
    for txbx_content in root.iter(qn('w:txbxContent')):
        # Récupérer tous les paragraphes (w:p) dans le contenu
        for p_elem in txbx_content.iter(qn('w:p')):
            try:
                # Créer un objet Paragraph python-docx à partir de l'élément XML
                yield Paragraph(p_elem, txbx_content)
            except Exception:
                # Si la création du Paragraph échoue, continuer
                pass

def apply_syllables(doc):
    """
    Parcourt le document (paragraphes, cellules de tableau, zones texte inline, zones texte VML)
    et remplace chaque mot par des runs colorés par syllabe.

    Méthode :
    - On extrait le texte du paragraphe, on vide le paragraphe (p.clear())
    - On parcourt caractère par caractère en segmentant les mots avec WORD_PATTERN
    - Pour chaque mot, on normalise et on appelle la fonction syllables(mot_norm)
      fournie par lirecouleur, puis on recrée des runs colorés en s'alignant
      sur la graphie originale (heuristique pour gérer les accents).
    """
    counter = [0]  # compteur partagé pour alterner les couleurs
    
    # Collecter tous les paragraphes à traiter
    paragraphs_to_process = []
    
    # Paragraphes du document principal
    paragraphs_to_process.extend(doc.paragraphs)
    
    # Inclure les cellules de tableaux
    for table in doc.tables:
        for cell in table._cells:
            paragraphs_to_process.extend(cell.paragraphs)

    # Inclure les zones de texte inline DrawingML (s'il y en a)
    for shape in doc.inline_shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            paragraphs_to_process.extend(shape.text_frame.paragraphs)
    
    # Inclure les zones de texte VML (legacy Word text boxes)
    paragraphs_to_process.extend(_iter_vml_textbox_paragraphs(doc))

    # Séparateurs considérés 'à part' (y compris tirets/traits) => on les rend tels quels
    SEPARATORS = set(" \t\n\r.,;:!?()[]{}\u00ab\u00bb\u201c\u201d'’/\\-–—*+=<>@#$%^&~")

    # Traiter tous les paragraphes collectés
    for p in paragraphs_to_process:
            if not p.text.strip():
                continue
            texte = p.text
            
            # Sauvegarder le formatage du premier run pour le préserver
            original_font_name = None
            original_font_size = None
            original_font_bold = None
            original_font_italic = None
            if len(p.runs) > 0:
                first_run = p.runs[0]
                original_font_name = first_run.font.name
                original_font_size = first_run.font.size
                original_font_bold = first_run.font.bold
                original_font_italic = first_run.font.italic
            
            p.clear()
            i = 0
            while i < len(texte):
                c = texte[i]
                # Conserver les séparateurs et ponctuations inchangés
                if c.isspace() or c in SEPARATORS:
                    run = p.add_run(c)
                    # Préserver le formatage original
                    if original_font_name:
                        try:
                            run.font.name = original_font_name
                        except Exception:
                            pass
                    if original_font_size:
                        try:
                            run.font.size = original_font_size
                        except Exception:
                            pass
                    if original_font_bold is not None:
                        try:
                            run.font.bold = original_font_bold
                        except Exception:
                            pass
                    if original_font_italic is not None:
                        try:
                            run.font.italic = original_font_italic
                        except Exception:
                            pass
                    i += 1
                    continue

                match = WORD_PATTERN.match(texte[i:])
                if match:
                    mot = match.group()
                    mot_norm = normalize(mot)

                    # Segmentation via lirecouleur.pylirecouleur
                    try:
                        syll_parts = syllables(mot_norm)
                    except Exception:
                        syll_parts = [mot]  # fallback: mot entier

                    # Recomposer en respectant la graphie d'origine (tentative)
                    # syll_parts est liste de syllabes (en graphie normalisée)
                    # On parcourt mot et on tente d'aligner chaque syllabe sur la chaîne originale.
                    pos = 0
                    for si, syl in enumerate(syll_parts):
                        # chercher syl dans mot (heuristique) à partir de pos
                        # on utilise une recherche insensible aux accents (on compare mot_norm)
                        target = syl
                        # calculer la longueur approximative en caractères dans mot
                        # on tente de consommer progressivement
                        # fallback : prendre une tranche égale en longueur
                        if si < len(syll_parts) - 1:
                            # essayer de trouver la correspondance exacte dans la portion restante
                            rem_norm = mot_norm[pos:]
                            idx = rem_norm.find(target)
                            if idx != -1:
                                real_idx = pos + idx
                                # map to original indices: approx by lengths
                                real_start = sum(len(x) for x in mot[:real_idx]) if False else pos  # best-effort
                            else:
                                # fallback heuristique
                                real_start = pos
                        else:
                            real_start = pos

                        # naive: découpage par longueur proportionnelle
                        # On prend pour la syllabe la même longueur que 'target' en caractères approximatifs
                        L = len(target)
                        part = mot[pos:pos+L] if pos+L <= len(mot) else mot[pos:]
                        # ajouter run coloré pour 'part'
                        run = p.add_run(part)
                        # appliquer couleur alternée
                        color = COUL_SYLL[counter[0] % len(COUL_SYLL)]
                        try:
                            run.font.color.rgb = color
                        except Exception:
                            pass
                        # Préserver le formatage original
                        if original_font_name:
                            try:
                                run.font.name = original_font_name
                            except Exception:
                                pass
                        if original_font_size:
                            try:
                                run.font.size = original_font_size
                            except Exception:
                                pass
                        if original_font_bold is not None:
                            try:
                                run.font.bold = original_font_bold
                            except Exception:
                                pass
                        if original_font_italic is not None:
                            try:
                                run.font.italic = original_font_italic
                            except Exception:
                                pass
                        counter[0] += 1
                        pos += len(part)

                    # si reste des caractères (rare), les ajouter
                    if pos < len(mot):
                        rest = mot[pos:]
                        run = p.add_run(rest)
                        try:
                            run.font.color.rgb = COUL_SYLL[counter[0] % len(COUL_SYLL)]
                        except Exception:
                            pass
                        # Préserver le formatage original
                        if original_font_name:
                            try:
                                run.font.name = original_font_name
                            except Exception:
                                pass
                        if original_font_size:
                            try:
                                run.font.size = original_font_size
                            except Exception:
                                pass
                        if original_font_bold is not None:
                            try:
                                run.font.bold = original_font_bold
                            except Exception:
                                pass
                        if original_font_italic is not None:
                            try:
                                run.font.italic = original_font_italic
                            except Exception:
                                pass
                        counter[0] += 1

                    i += len(mot)
                else:
                    # caractère isolé (non matched by WORD_PATTERN) => ajouter tel quel
                    run = p.add_run(texte[i])
                    # Préserver le formatage original
                    if original_font_name:
                        try:
                            run.font.name = original_font_name
                        except Exception:
                            pass
                    if original_font_size:
                        try:
                            run.font.size = original_font_size
                        except Exception:
                            pass
                    if original_font_bold is not None:
                        try:
                            run.font.bold = original_font_bold
                        except Exception:
                            pass
                    if original_font_italic is not None:
                        try:
                            run.font.italic = original_font_italic
                        except Exception:
                            pass
                    i += 1