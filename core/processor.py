#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
#
# core/processor.py – Orchestrateur central (appelé par main.py)
#
# Rôle :
# - Être le point d'entrée principal pour le traitement d'un fichier .docx
#   (main.py appelle process_document).
# - Orchestrer les étapes : mise en forme de base, coloration syllabique,
#   grisement des lettres muettes, coloration des nombres, sauvegarde et ouverture.
#
# Conception :
# - Les fonctions spécialisées sont dans :
#     core.formatter.apply_base_formatting
#     core.syllables.apply_syllables
#     core.mute_letters.apply_mute_letters
#     core.numbers_position.apply_position_numbers
#     core.numbers_multicolor.apply_multicolor_numbers
# - processor.py coordonne l'ordre d'application et gère les fichiers temporaires
#   si nécessaire (pour éviter que les traitements n'interfèrent entre eux).
#
# Compatibilité :
# - La signature process_document(filepath, progress_callback=None, open_after=True)
#   reste compatible avec les appels existants (open_after=True par défaut).
#

import os
import platform
import subprocess
import tempfile
import traceback

from docx import Document
from typing import Optional

from config.settings import options
from ui.interface import update_progress

# Import des fonctions spécialisées (formatter expose apply_base_formatting)
from .formatter import apply_base_formatting
from .syllables import apply_syllables
from .mute_letters import apply_mute_letters
# Nouvelle API combinée (remplace complètement l'ancienne logique _apply_both_with_temp)
from .syllables_mute import apply_syllables_mute

# Coloration des nombres
from .numbers_position import apply_position_numbers
from .numbers_multicolor import apply_multicolor_numbers


def _save_output_and_open(doc: Document, input_filepath: str, open_after: bool = True) -> str:
    """
    Sauvegarde le document dans un sous-dossier 'DYS' à côté du fichier d'entrée,
    en évitant d'écraser un fichier existant (ajoute (n) si besoin).

    Si open_after est True : ouvre le fichier final avec l'application native du système.
    Si open_after est False : ne lance pas l'application (permet un traitement batch).

    Retourne le chemin complet du fichier de sortie (.docx).
    """
    # Calcul du dossier de sortie et nom de fichier
    folder = os.path.join(os.path.dirname(input_filepath), "DYS")
    os.makedirs(folder, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_filepath))[0]
    output = os.path.join(folder, f"{base}_DYS.docx")
    i = 1
    while os.path.exists(output):
        output = os.path.join(folder, f"{base}_DYS ({i}).docx")
        i += 1

    # Sauvegarde
    doc.save(output)
    update_progress(100, f"Terminé → {os.path.basename(output)}")

    # Ouverture automatique (si demandée)
    if open_after:
        try:
            if platform.system() == "Linux":
                subprocess.call(["xdg-open", output])
            elif platform.system() == "Darwin":
                subprocess.call(["open", output])
            else:
                # Windows
                os.startfile(output)
        except Exception:
            # Ne pas planter l'application si l'ouverture échoue, on signale simplement
            update_progress(100, f"Sauvegardé → {os.path.basename(output)} (ouverture automatique impossible)")
    else:
        # Mode batch : on n'ouvre pas le fichier
        pass

    return output


def process_document(filepath: str, progress_callback=None, open_after: bool = True) -> str:
    """
    Point d'entrée principal pour le traitement d'un .docx.
    - Ouvre le fichier
    - Applique la mise en forme de base via apply_base_formatting
    - Applique ensuite syllabes et/ou muettes selon les options utilisateur
    - Applique la coloration des nombres selon les options utilisateur
    - Sauvegarde le résultat et l'ouvre (si open_after=True)

    Retourne le chemin du fichier de sortie (.docx).

    Notes :
    - open_after permet de désactiver l'ouverture automatique (utile en traitement
      batch où l'on veut ouvrir le dossier final à la fin).
    - progress_callback est ignoré par défaut (update_progress global est utilisé).
    """
    # Début : informer l'UI que le traitement commence
    update_progress(10, "Ouverture du document...")
    try:
        doc = Document(filepath)
    except Exception as e:
        update_progress(0, "Échec ouverture document")
        raise

    # Récupération des options de l'UI (config.settings.options fournit les widgets)
    police = options['police'].get()
    taille_pt = options['taille'].get()
    espacement = options['espacement'].get()
    interlignes = options['interligne'].get()
    syllabes_on = options['syllabes'].get()
    muettes_on = options['griser_muettes'].get()

    # --- Étape 1 : traitements spécialisés (syllabes / muettes) - AVANT mise en forme ---
    try:
        # Cas : si seul syllabes activé
        if syllabes_on and not muettes_on:
            update_progress(20, "Coloration syllabique...")
            apply_syllables(doc)

        # Cas : si seules muettes activées
        elif muettes_on and not syllabes_on:
            update_progress(20, "Grisage des lettres muettes...")
            apply_mute_letters(doc)

        # Cas : les deux activés -> utiliser la nouvelle fonction dédiée apply_syllables_mute
        elif syllabes_on and muettes_on:
            update_progress(20, "Coloration syllabique + lettres muettes...")
            # apply_syllables_mute retourne un Document résultant (avec syllabes+muettes appliqués)
            doc = apply_syllables_mute(doc, filepath)

        # Cas : aucun traitement spécial
        else:
            update_progress(20, "Aucun traitement syllabique/muettes demandé")
    except Exception:
        # Log minimal et propagation (main.py / UI doit afficher l'erreur)
        traceback.print_exc()
        update_progress(0, "Erreur durant les traitements spécialisés")
        raise

    # --- Étape 2 : coloration des nombres - AVANT mise en forme ---
    try:
        # options['position'] et options['multicolore'] sont des tk.Variable dans l'UI.
        # On est défensif : utiliser options.get() puis .get() si présent.
        num_pos_var = options.get('position', None)
        num_multi_var = options.get('multicolore', None)
        pos_val = num_pos_var.get() if num_pos_var is not None else False
        multi_val = num_multi_var.get() if num_multi_var is not None else False

        # L'UI synchronise déjà les deux cases (mutuellement exclusives), mais on gère tous les cas.
        if multi_val and not pos_val:
            update_progress(40, "Coloration multicolore des nombres...")
            apply_multicolor_numbers(doc)
        elif pos_val and not multi_val:
            update_progress(40, "Coloration par position des nombres...")
            apply_position_numbers(doc)
        elif pos_val and multi_val:
            # cas improbable puisque l'UI empêche normalement les deux cochés,
            # donner priorité à multicolore pour éviter comportement inattendu.
            update_progress(40, "Coloration multicolore des nombres (priorité multicolore)...")
            apply_multicolor_numbers(doc)
        else:
            # aucune option nombres activée
            update_progress(40, "Aucune coloration numérique demandée")
    except Exception:
        traceback.print_exc()
        update_progress(0, "Erreur coloration nombres")
        raise

    # --- Étape 3 : mise en forme de base FINALE (police, taille, espacement, interligne) ---
    # Appliquée EN DERNIER pour s'appliquer sur tous les runs colorés
    try:
        update_progress(70, "Application finale de la mise en forme...")
        apply_base_formatting(doc, police, taille_pt, espacement, interlignes)
    except Exception:
        update_progress(0, "Échec mise en forme finale")
        raise

    # --- Étape finale : sauvegarde et ouverture conditionnelle ---
    try:
        update_progress(90, "Sauvegarde...")
        output_path = _save_output_and_open(doc, filepath, open_after=open_after)
    except Exception:
        update_progress(0, "Échec sauvegarde / ouverture")
        raise

    return output_path

# Fin du module