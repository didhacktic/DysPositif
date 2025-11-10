# -----------------
# core/formatter.py 
# -----------------
import os
import platform
import subprocess
from docx import Document
from docx.shared import Mm

from config.settings import options
from ui.interface import update_progress

from .syllables import apply_syllables
from .mute_letters import apply_mute_letters
from .numbers_multicolor import apply_multicolor_numbers
from .numbers_position import apply_position_numbers
from .a3_enlarger import apply_a3_format
from .utils import apply_font_consistently, apply_spacing_and_line_spacing


def format_document(filepath: str):
    update_progress(10, "Ouverture du document...")
    doc = Document(filepath)

    police = options['police'].get()
    taille_pt = options['taille'].get()
    espacement = options['espacement'].get()
    interlignes = options['interligne'].get()
    syllabes_on = options['syllabes'].get()
    muettes_on = options['griser_muettes'].get()
    multicolore_on = options['multicolore'].get()
    position_on = options['position'].get()
    format_a3 = options['format'].get() == "A3"
    agrandir_objets = options['agrandir_objets'].get()

    update_progress(20, "Police + taille...")
    apply_font_consistently(doc, police, taille_pt)
    apply_spacing_and_line_spacing(doc, espacement, interlignes)

    # SYLLABES D'ABORD
    if syllabes_on:
        update_progress(40, "Coloration syllabique rouge/bleu...")
        apply_syllables(doc)

    # MUETTES ENSUITE (par-dessus, sans écraser la couleur)
    if muettes_on:
        update_progress(60, "Grisage des lettres muettes (sans écraser la couleur)...")
        apply_mute_letters(doc)

    # NOMBRES
    if multicolore_on:
        update_progress(70, "Coloration multicolore...")
        apply_multicolor_numbers(doc)
    if position_on:
        update_progress(70, "Coloration par position...")
        apply_position_numbers(doc)

    # A3
    if format_a3:
        update_progress(80, "Format A3...")
        apply_a3_format(doc, agrandir_objets)

    # SAUVEGARDE
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
    update_progress(100, f"Terminé ! → DYS/{os.path.basename(output)}")

    sys = platform.system()
    if sys == "Linux":
        subprocess.call(["xdg-open", output])
    elif sys == "Darwin":
        subprocess.call(["open", output])
    else:
        os.startfile(output)
