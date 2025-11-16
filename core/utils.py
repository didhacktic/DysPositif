# core/utils.py
"""
Utilitaires sûrs pour modifications fines sur docx (python-docx).

Fonctions principales exposées :
- apply_font_consistently(doc, police_name, taille_pt_value, include_shapes=False, include_tables=False)
    -> applique nom et taille de police SUR LE CORPS PRINCIPAL par défaut;
       si include_shapes=True, applique aussi aux text_frames (shapes) (option heuristique).
    -> ne modifie PAS run.font.spacing (tracking/kerning).
    -> ne touche PAS headers/footers par défaut.

- apply_spacing_and_line_spacing(doc, espacement: bool, interlignes: bool)
    -> wrapper de compatibilité : n'applique pas le tracking (espacement lettres),
       mais applique l'interlignage si demandé.

- split_run_and_color(run, start, end, color_hex) et safe_color_substring_in_paragraph(...)
    -> opérations sûres pour colorer une sous-portion d'un run sans insérer d'espaces.

- adjust_textboxes_after_font_change(doc, scale_step=1.12, max_iter=5, enable_word_wrap=True)
    -> heuristique pour réduire le clipping dans les text boxes après changement de police.
"""
from __future__ import annotations
from copy import deepcopy
from typing import Optional, Iterable, Tuple
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx

# -----------------------
# Helpers généraux
# -----------------------
def _is_body_paragraph(paragraph) -> bool:
    """
    True si paragraph appartient au corps principal (document), et non header/footer/cell/shape.
    """
    # paragraph._p est l'élément XML (w:p)
    parent = paragraph._p.getparent()
    if parent is None:
        return False
    # Exclure si parent est une cellule de tableau (w:tc)
    try:
        if parent.tag.endswith('}tc'):
            return False
    except Exception:
        pass
    # Remonter pour détecter hdr/ftr
    el = paragraph._p
    while el is not None:
        tag = getattr(el, 'tag', '')
        if tag.endswith('}hdr') or tag.endswith('}ftr') or tag.endswith('}header') or tag.endswith('}footer'):
            return False
        el = el.getparent()
    return True

def _iter_body_paragraphs(doc: Document):
    """
    Itère uniquement les paragraphes du corps principal (exclut headers/footers/tables/zones graphiques).
    """
    for p in doc.paragraphs:
        if _is_body_paragraph(p):
            yield p

def _iter_table_paragraphs(doc: Document):
    """
    Itère les paragraphes situés dans les cellules de tableaux (tableau optionnel).
    """
    for table in getattr(doc, "tables", ()):
        for cell in getattr(table, "_cells", ()):
            for p in getattr(cell, "paragraphs", ()):
                yield p

def _iter_shape_text_frames(doc: Document):
    """
    Itère (shape, text_frame) pour les formes accessibles via python-docx.
    Robuste : certaines installations n'exposent pas les mêmes wrappers.
    """
    # inline_shapes (souvent text boxes/icônes inline)
    try:
        for shape in getattr(doc, "inline_shapes", ()):
            tf = getattr(shape, "text_frame", None)
            if tf is not None:
                yield shape, tf
    except Exception:
        pass

    # attempt shapes collection (floating shapes) si exposée
    try:
        for shape in getattr(doc, "shapes", ()):
            tf = getattr(shape, "text_frame", None)
            if tf is not None:
                yield shape, tf
    except Exception:
        pass

# -----------------------
# Fonctions principales
# -----------------------
def apply_font_consistently(doc: Document, police_name: str, taille_pt_value: int,
                            include_shapes: bool = False, include_tables: bool = False):
    """
    Applique le nom de police et la taille aux runs du CORPS PRINCIPAL uniquement par défaut.
    - include_shapes=False : ne touche pas les text boxes / shapes
    - include_tables=False : ne touche pas les paragraphes des cellules de tableaux
    - NE MODIFIE PAS run.font.spacing (tracking/kerning)
    - Préserve autres propriétés de paragraphes (alignment, keep_with_next, page_break_before...)
    """
    taille = Pt(taille_pt_value)

    # Corps principal
    for p in _iter_body_paragraphs(doc):
        for run in p.runs:
            try:
                run.font.name = police_name
            except Exception:
                pass
            try:
                run.font.size = taille
            except Exception:
                pass
        # tenter de mettre à jour le style du paragraphe sans toucher les autres propriétés
        try:
            if p.style and hasattr(p.style, 'font'):
                p.style.font.name = police_name
                p.style.font.size = taille
        except Exception:
            pass

    # Option : appliquer dans les tableaux (cellules) si explicitement demandé
    if include_tables:
        for p in _iter_table_paragraphs(doc):
            for run in p.runs:
                try:
                    run.font.name = police_name
                except Exception:
                    pass
                try:
                    run.font.size = taille
                except Exception:
                    pass

    # Option : appliquer dans les shapes/text_frames si demandé
    if include_shapes:
        for shape, tf in _iter_shape_text_frames(doc):
            try:
                # text_frame.paragraphs similaire à document.paragraphs (si implémenté)
                for p in getattr(tf, "paragraphs", ()):
                    for run in getattr(p, "runs", ()):
                        try:
                            run.font.name = police_name
                        except Exception:
                            pass
                        try:
                            run.font.size = taille
                        except Exception:
                            pass
            except Exception:
                # fallback : shape may expose text via shape.text
                try:
                    text = getattr(shape, "text", None)
                    if text is not None:
                        # cannot set runs easily; skip to be safe
                        pass
                except Exception:
                    pass

def apply_line_spacing(doc: Document, interlignes_value: Optional[float] = None):
    """
    Applique l'interlignage (line spacing multiplier) aux paragraphes du corps principal uniquement.
    """
    if interlignes_value is None:
        return
    for p in _iter_body_paragraphs(doc):
        try:
            p.paragraph_format.line_spacing = interlignes_value
        except Exception:
            pass

def apply_spacing_and_line_spacing(doc: Document, espacement: bool, interlignes: bool):
    """
    Wrapper de compatibilité pour l'ancienne API.
    - espacement : bool (anciennement activait run.font.spacing) => IGNORÉ par sécurité pour éviter
                  l'espacement excessif.
    - interlignes : bool => applique interlignage 1.5 si True.
    """
    # Nous n'activons pas run.font.spacing (tracking) pour éviter l'effet très espacé.
    if interlignes:
        apply_line_spacing(doc, 1.5)

# -----------------------
# Manipulation sûre de runs (split / clone / color)
# -----------------------
def _clone_run_element_with_text(run, text: str):
    """
    Clone l'élément XML d'un run (run._r) en conservant les propriétés et remplace le texte par `text`.
    """
    new_r = deepcopy(run._r)
    for t in new_r.iterfind('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        t.text = text
    return new_r

def _insert_run_after(run, new_r_element):
    """Insère new_r_element (élément XML) juste après run._r"""
    run._r.addnext(new_r_element)

def _set_color_on_run_element(run_element, color_hex: str):
    """Ajoute/modifie la couleur (w:color) sur un élément run (XML)."""
    rpr = run_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        run_element.insert(0, rpr)
    color_el = OxmlElement('w:color')
    color_el.set(qn('w:val'), color_hex)
    # supprimer éventuelles couleurs existantes
    for child in list(rpr):
        if child.tag.endswith('}color'):
            rpr.remove(child)
    rpr.append(color_el)

def split_run_and_color(run, start: int, end: int, color_hex: str) -> None:
    """
    Découpe le run `run` en head|middle|tail et applique color_hex uniquement sur middle.
    Préserve les propriétés rPr du run original (deepcopy).
    start inclusif, end exclusif (indices Python).
    """
    text = run.text or ""
    if not text:
        return
    if start < 0:
        start = 0
    if end > len(text):
        end = len(text)
    if start >= end:
        return

    head = text[:start]
    middle = text[start:end]
    tail = text[end:]

    # remplacer le run original par head
    run.text = head

    middle_el = _clone_run_element_with_text(run, middle)
    tail_el = _clone_run_element_with_text(run, tail)

    _insert_run_after(run, middle_el)
    # middle_el est un element; on insère tail après middle
    middle_el.addnext(tail_el)

    _set_color_on_run_element(middle_el, color_hex)

    if tail == "":
        try:
            tail_el.getparent().remove(tail_el)
        except Exception:
            pass

def safe_color_substring_in_paragraph(paragraph, substring: str, color_hex: str, first_only: bool = False):
    """
    Colorer la première (ou toutes) occurrences d'une sous-chaîne contenue DANS UN SEUL run.
    Limitation: n'interprète pas les occurrences traversant plusieurs runs.
    """
    if not substring:
        return
    for run in paragraph.runs:
        t = run.text or ""
        idx = t.find(substring)
        if idx >= 0:
            split_run_and_color(run, idx, idx + len(substring), color_hex)
            if first_only:
                return

# -----------------------
# Ajustement heuristique des text boxes (shapes)
# -----------------------
def _iter_all_shapes(doc: Document):
    """
    Itère sur les objets shape accessibles (inline_shapes et shapes si exposés).
    """
    try:
        for s in getattr(doc, "inline_shapes", ()):
            yield s
    except Exception:
        pass
    try:
        for s in getattr(doc, "shapes", ()):
            yield s
    except Exception:
        pass

def adjust_textboxes_after_font_change(doc: Document, scale_step: float = 1.12, max_iter: int = 5, enable_word_wrap: bool = True):
    """
    Heuristique pour réduire le clipping dans les text boxes après changement de police :
      - tente d'activer word_wrap sur text_frame si exposé,
      - agrandit height/width progressivement selon scale_step et max_iter.
    Usage : appeler APRES avoir modifié la police dans les shapes (si include_shapes=True).
    Attention : heuristique ; pas de garantie parfaite. Toujours tester sur une copie du .docx.
    """
    for shape in _iter_all_shapes(doc):
        try:
            tf = getattr(shape, "text_frame", None)
            if tf is None:
                continue

            # activer word_wrap si disponible
            if enable_word_wrap:
                try:
                    tf.word_wrap = True
                except Exception:
                    pass

            # obtenir height/width si exposés (EMUs)
            height = getattr(shape, "height", None)
            width = getattr(shape, "width", None)
            if height is None:
                continue

            # itération d'agrandissement
            for i in range(max_iter):
                try:
                    # augmenter progressivement la hauteur; élargir légèrement la largeur
                    new_height = int(height * (scale_step ** (i + 1)))
                    shape.height = new_height
                    if width is not None:
                        # élargir légèrement pour le cas d'enroulement horizontal
                        shape.width = int(width * (1.0 + 0.08 * (i + 1)))
                except Exception:
                    break
        except Exception:
            continue

# -----------------------
# Fin du module
# -----------------------