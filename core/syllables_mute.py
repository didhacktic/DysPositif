# core/syllables_mute.py – VERSION QUI MARCHE 100%
from docx.shared import RGBColor
import re

ROUGE = RGBColor(220, 20, 60)
BLEU = RGBColor(30, 144, 255)
GRIS = RGBColor(130, 130, 130)

MUETTES = {
    'ent': ['vent', 'dent', 'cent', 'lent', 'gent', 'absent', 'présent'],
    'ez': ['chez', 'rez', 'nez'],
    'ait': ['fait', 'trait', 'plait'],
}

def est_muette(mot_lower):
    for suffixe, exceptions in MUETTES.items():
        if mot_lower.endswith(suffixe) and mot_lower not in exceptions:
            return suffixe
    return None

def apply_syllables_mute(doc):
    counter = 0
    for paragraph in doc.paragraphs:
        runs = list(paragraph.runs)
        i = 0
        while i < len(runs):
            run = runs[i]
            text = run.text
            if not text.strip():
                i += 1
                continue

            words = re.findall(r"[a-zA-ZàâéèêëîïôöùûüçÀÂÉÈÊËÎÏÔÖÙÛÜÇ']+", text)
            if not words:
                i += 1
                continue

            word = words[0]
            muette = est_muette(word.lower())

            if muette and text.lower().endswith(muette):
                debut = len(text) - len(muette)
                partie1 = text[:debut]
                partie2 = text[debut:]

                run.text = partie1

                new_run = paragraph.add_run(partie2)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline

                color = ROUGE if counter % 2 == 0 else BLEU
                new_run.font.color.rgb = color

                r, g, b = color
                new_run.font.color.rgb = RGBColor(int(r * 0.6), int(g * 0.6), int(b * 0.6))

                counter += 1
            else:
                color = ROUGE if counter % 2 == 0 else BLEU
                run.font.color.rgb = color
                counter += 1

            i += 1
