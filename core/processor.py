#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
#
# core/processor.py – Orchestrateur central (appelé par main.py)
#
# Rôle :
# - Être le point d'entrée principal pour le traitement d'un fichier .docx
#   (main.py appelle process_document).
# - Orchestrer les étapes : mise en forme de base, coloration syllabique,
#   grisement des lettres muettes, sauvegarde et ouverture.
#
# Conception :
# - Les fonctions spécialisées sont dans :
#     core.formatter.apply_base_formatting
#     core.syllables.apply_syllables
#     core.mute_letters.apply_mute_letters
# - processor.py coordonne l'ordre d'application et gère les fichiers temporaires
#   si nécessaire (pour éviter que les traitements n'interfèrent entre eux).
#
# Compatibilité :
# - Conserve la signature attendue par main.py : process_document(filepath, progress_callback=None)
#   (progress_callback est ignoré car update_progress est global).
#

import os
import platform
import subprocess
import tempfile
import traceback

from docx import Document

from config.settings import options
from ui.interface import update_progress

# Import des fonctions spécialisées (formatter expose apply_base_formatting)
from .formatter import apply_base_formatting
from .syllables import apply_syllables
from .mute_letters import apply_mute_letters
# Nouvelle API combinée (remplace complètement l'ancienne logique _apply_both_with_temp)
from .syllables_mute import apply_syllables_mute


def _save_output_and_open(doc: Document, input_filepath: str):
    """
    Sauvegarde le document dans un sous-dossier 'DYS' à côté du fichier d'entrée,
    en évitant d'écraser un fichier existant (ajoute (n) si besoin), puis ouvre le
    fichier final avec l'application native du système.
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

    # Ouverture du fichier final selon l'OS
    try:
        if platform.system() == "Linux":
            subprocess.call(["xdg-open", output])
        elif platform.system() == "Darwin":
            subprocess.call(["open", output])
        else:
            os.startfile(output)
    except Exception:
        # Ne pas planter l'application si l'ouverture échoue
        update_progress(100, f"Sauvegardé → {os.path.basename(output)} (ouverture automatique impossible)")
        return


def process_document(filepath: str, progress_callback=None):
    """
    Point d'entrée principal pour le traitement d'un .docx.
    - Ouvre le fichier
    - Applique la mise en forme de base via apply_base_formatting
    - Applique ensuite syllabes et/ou muettes selon les options utilisateur
    - Sauvegarde le résultat et l'ouvre

    Signature compatible avec main.py : process_document(filepath, progress_callback=None)
    (progress_callback est ignoré au profit de update_progress global).
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

    # --- Étape 1 : mise en forme de base (police, taille, espacement, interligne) ---
    try:
        update_progress(20, "Application de la mise en forme de base...")
        apply_base_formatting(doc, police, taille_pt, espacement, interlignes)
    except Exception:
        update_progress(0, "Échec mise en forme de base")
        # On continue ? Ici on choisit de propager pour que l'UI affiche l'erreur.
        raise

    # --- Étape 2 : traitements spécialisés (syllabes / muettes) ---
    try:
        # Cas : si seul syllabes activé
        if syllabes_on and not muettes_on:
            update_progress(50, "Coloration syllabique...")
            apply_syllables(doc)

        # Cas : si seules muettes activées
        elif muettes_on and not syllabes_on:
            update_progress(50, "Grisage des lettres muettes...")
            apply_mute_letters(doc)

        # Cas : les deux activés -> utiliser la nouvelle fonction dédiée apply_syllables_mute
        elif syllabes_on and muettes_on:
            # apply_syllables_mute retourne un Document résultant (avec syllabes+muettes appliqués)
            doc = apply_syllables_mute(doc, filepath)

        # Cas : aucun traitement spécial -> on garde seulement la mise en forme de base
        else:
            update_progress(50, "Aucun traitement syllabique/muettes demandé")
    except Exception:
        # Log minimal et propagation (main.py / UI doit afficher l'erreur)
        traceback.print_exc()
        update_progress(0, "Erreur durant les traitements spécialisés")
        raise

    # --- Étape finale : sauvegarde et ouverture ---
    try:
        update_progress(90, "Sauvegarde...")
        _save_output_and_open(doc, filepath)
    except Exception:
        update_progress(0, "Échec sauvegarde / ouverture")
        raise

# Fin du module
