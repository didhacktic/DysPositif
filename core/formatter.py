# core/formatter.py
import os
import platform
import subprocess
import tempfile
from docx import Document

from config.settings import options
from ui.interface import update_progress

from .syllables import apply_syllables
from .mute_letters import apply_mute_letters
from .utils import apply_font_consistently, apply_spacing_and_line_spacing

def format_document(filepath: str):
    update_progress(10, "Ouverture...")
    doc = Document(filepath)

    police = options['police'].get()
    taille_pt = options['taille'].get()
    espacement = options['espacement'].get()
    interlignes = options['interligne'].get()
    syllabes_on = options['syllabes'].get()
    muettes_on = options['griser_muettes'].get()

    apply_font_consistently(doc, police, taille_pt)
    apply_spacing_and_line_spacing(doc, espacement, interlignes)

    # --- SYLLABES SEULES ---
    if syllabes_on and not muettes_on:
        update_progress(50, "Coloration syllabique...")
        apply_syllables(doc)

    # --- MUETTES SEULES ---
    elif muettes_on and not syllabes_on:
        update_progress(50, "Grisage muettes...")
        apply_mute_letters(doc)

    # --- LES DEUX : SYLLABES → TMP → MUETTES ---
    elif syllabes_on and muettes_on:
        update_progress(50, "Étape 1/2 : Syllabes...")
        temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
        os.close(temp_fd)
        doc.save(temp_path)

        # Appliquer syllabes sur doc original
        apply_syllables(doc)

        # Charger le doc modifié et appliquer muettes
        update_progress(75, "Étape 2/2 : Muettes...")
        temp_doc = Document(temp_path)
        apply_mute_letters(temp_doc)
        temp_doc.save(temp_path)

        # Remplacer le doc original
        doc._element.body._element = temp_doc._element.body._element

        # Nettoyage
        os.unlink(temp_path)

    # --- SAUVEGARDE ---
    update_progress(90, "Sauvegarde...")
    dossier_dys = os.path.join(os.path.dirname(filepath), "DYS")
    os.makedirs(dossier_dys, exist_ok=True)
    base = os.path.splitext(os.path.basename(filepath))[0]
    output = os.path.join(dossier_dys, f"{base}_DYS.docx")
    i = 1
    while os.path.exists(output):
        output = os.path.join(dossier_dys, f"{base}_DYS ({i}).docx")
        i += 1

    doc.save(output)
    update_progress(100, f"Terminé → {os.path.basename(output)}")

    if platform.system() == "Linux":
        subprocess.call(["xdg-open", output])
    elif platform.system() == "Darwin":
        subprocess.call(["open", output])
    else:
        os.startfile(output)