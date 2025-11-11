# core/mute_letters.py – Grisage des muettes (seul) – TS TOUJOURS GRISÉS
import re
from docx.shared import RGBColor

GRIS = RGBColor(180, 180, 180)

# Exceptions globales
EXCEPTIONS_GLOBALES = {
    'le', 'me', 'te', 'se', 'ce', 'ne', 'de', 'je', 'que', 'être',
    'est', 'et', 'tous', 'des', 'les', 'tes', 'ces', 'ses', 'mes',
    'une', 'la', 'ma', 'ta', 'sa', 'du', 'au', 'aux', 'un'
}

# Exceptions spécifiques
EXCEPTIONS_SUFFIXE = {
    'ent': ['vent', 'dent', 'cent', 'lent', 'gent', 'absent', 'présent'],
    'x': ['lux']
}

ENT_T_SEUL = ['argent', 'agent', 'urgent']

WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ'’-]+")
H_MUET_PATTERN = re.compile(r"\b[hH][a-zA-ZàâéèêëîïôöùûüçÀÂÉÈÊËÎÏÔÖÙÛÜÇ']*", re.UNICODE)

def est_nom_propre(mot: str) -> bool:
    if not mot:
        return False
    return mot[0].isupper() and mot.lower() not in {'il', 'elle', 'on', 'nous', 'vous', 'ils', 'elles'}

def est_adverbe_ment(mot_lower: str) -> bool:
    return mot_lower.endswith('ment')

def est_participe(mot_lower: str) -> bool:
    return mot_lower.endswith(('éent', 'ient', 'uent'))

def est_muette_finale(mot: str, mot_lower: str) -> str:
    if not mot:
        return None
    if est_nom_propre(mot):
        return None
    if mot_lower in EXCEPTIONS_GLOBALES:
        return None
    if mot_lower.endswith('cts') or mot_lower.endswith('ct'):
        return None
    if mot_lower.endswith('ées'):
        return 'es'

    if mot_lower.endswith('ents') or mot_lower.endswith('ent'):
        base = mot_lower[:-1] if mot_lower.endswith('s') else mot_lower
        if mot_lower in ENT_T_SEUL or base in ENT_T_SEUL:
            return 't'
        if est_adverbe_ment(base) or est_participe(base):
            return 't'
        if base.endswith('c') or base.endswith('g'):
            return 'nts' if mot_lower.endswith('s') else 'nt'
        if base not in EXCEPTIONS_SUFFIXE['ent']:
            return 'ents' if mot_lower.endswith('s') else 'ent'

    if mot_lower.endswith('x') and mot_lower not in EXCEPTIONS_SUFFIXE['x']:
        return 'x'

    if mot_lower == 'plus':
        return 's'

    # ORDRE CRITIQUE : ts/ds AVANT s
    if mot_lower.endswith('ts'):
        return 'ts'
    if mot_lower.endswith('ds'):
        return 'ds'
    if mot_lower.endswith('t'):
        return 't'
    if mot_lower.endswith('d'):
        return 'd'
    if mot_lower.endswith('s'):
        return 's'

    if mot_lower.endswith('e'):
        if mot_lower.endswith('ce') or mot_lower.endswith('ge'):
            return None
        if not mot_lower.endswith('es'):
            return 'e'

    return None

def apply_mute_letters(doc):
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue

        original_style = None
        if paragraph.runs:
            first_run = paragraph.runs[0]
            original_style = {
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        style = paragraph.style
        alignment = paragraph.alignment
        full_text = paragraph.text
        paragraph.clear()

        i = 0
        while i < len(full_text):
            c = full_text[i]

            if not WORD_PATTERN.match(c):
                run = paragraph.add_run(c)
                if original_style:
                    run.bold = original_style['bold']
                    run.italic = original_style['italic']
                    run.underline = original_style['underline']
                i += 1
                continue

            match = WORD_PATTERN.search(full_text[i:])
            if not match:
                run = paragraph.add_run(c)
                if original_style:
                    run.bold = original_style['bold']
                    run.italic = original_style['italic']
                    run.underline = original_style['underline']
                i += 1
                continue

            word = match.group()
            i += len(word)
            word_lower = word.lower()

            # H muet : griser le "h", puis traiter le reste
            reste = word
            if H_MUET_PATTERN.match(word):
                h_part = word[0]
                reste = word[1:]
                run_h = paragraph.add_run(h_part)
                run_h.font.color.rgb = GRIS
                if original_style:
                    run_h.bold = original_style['bold']
                    run_h.italic = original_style['italic']
                    run_h.underline = original_style['underline']
                if not reste:
                    continue

            # Muette finale sur le reste
            muette = est_muette_finale(reste, reste.lower())
            if muette:
                debut = len(reste) - len(muette)
                partie1 = reste[:debut]
                partie2 = reste[debut:]

                if partie1:
                    run1 = paragraph.add_run(partie1)
                    if original_style:
                        run1.bold = original_style['bold']
                        run1.italic = original_style['italic']
                        run1.underline = original_style['underline']
                if partie2:
                    run2 = paragraph.add_run(partie2)
                    run2.font.color.rgb = GRIS
                    if original_style:
                        run2.bold = original_style['bold']
                        run2.italic = original_style['italic']
                        run2.underline = original_style['underline']
            else:
                run = paragraph.add_run(reste)
                if original_style:
                    run.bold = original_style['bold']
                    run.italic = original_style['italic']
                    run.underline = original_style['underline']

        paragraph.style = style
        if alignment:
            paragraph.alignment = alignment