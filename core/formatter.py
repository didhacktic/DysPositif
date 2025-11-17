"""
# core/formatter.py
# -------------------------------------------------
# Module responsable des opérations de mise en forme « de base » sur un
# document python-docx. Après la réorganisation, ce fichier n'orchestration
# plus l'ensemble du pipeline : il expose des fonctions ciblées que
# processor.py appellera.
#
# Fonctions exposées :
# - apply_base_formatting(doc, police, taille_pt, espacement, interlignes)
#     applique la police, la taille, l'espacement des caractères et
#     l'interlignage de manière cohérente sur l'ensemble du document.
#
# NOTES :
# - update_progress reste une responsabilité du caller (processor.py)
# - les traitements plus spécialisés (syllabes / muettes) sont dans
#   core.syllables et core.mute_letters et doivent être appelés depuis
#   processor.py selon la logique souhaitée.
"""

import os
from docx import Document

from config.settings import options
from ui.interface import update_progress

# Utilitaires qui effectuent les modifications fines sur le document
from .utils import apply_font_consistently, apply_spacing_and_line_spacing


def apply_base_formatting(doc: Document, police: str, taille_pt: int, espacement: float, interlignes: float):
    """Applique la mise en forme de base à un document python-docx.

    Arguments :
        doc : instance docx.Document à modifier (modifié in-place)
        police : nom de la police à appliquer
        taille_pt : taille de la police en points
        espacement : espacement des caractères (tracking/kerning)
        interlignes : valeur d'interligne désirée

    La fonction modifie le document en place et ne retourne rien.
    """
    # Mise à jour de l'UI (facultative mais utile pour le feedback utilisateur)
    try:
        update_progress(20, "Application police et taille...")
    except Exception:
        # Si update_progress n'est pas disponible pour une raison quelconque,
        # on continue silencieusement.
        pass

    # Appliquer la police et la taille de manière cohérente sur l'ensemble du document
    # Étendre la portée : inclure tableaux + headers/footers + shapes (DrawingML et VML)
    apply_font_consistently(doc, police, taille_pt, include_tables=True, include_headers_footers=True, include_shapes=True)

    # Appliquer l'espacement et l'interlignage
    apply_spacing_and_line_spacing(doc, espacement, interlignes)

    # Fin de la mise en forme de base
    try:
        update_progress(30, "Mise en forme de base appliquée")
    except Exception:
        pass


# Pour compatibilité ascendante, conserver une fonction minimaliste format_document
# qui applique uniquement la mise en forme de base (les traitements avancés
# doivent être orchestrés depuis processor.py)
def format_document(filepath: str):
    """Ancienne API (compat) : ouvre un document et applique la mise en forme de base.
    Note : cette fonction n'implémente plus le pipeline complet ; utilisez
    core.processor.process_document pour l'orchestration complète.
    """
    doc = Document(filepath)
    police = options['police'].get()
    taille_pt = options['taille'].get()
    espacement = options['espacement'].get()
    interlignes = options['interligne'].get()
    apply_base_formatting(doc, police, taille_pt, espacement, interlignes)
    return doc