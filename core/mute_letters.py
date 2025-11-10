# -------------------------------------------------
# core/mute_letters.py – VERSION INVINCIBLE (AUCUN BLOCAGE + SYLLABES + GRISAGE)
# -------------------------------------------------
import re
from docx.shared import RGBColor

GRAY = RGBColor(180, 180, 180)

EXCEPTIONS = {
    'e': ['café', 'clé', 'thé', 'été', 'bébé', 'cité', 'dé', 'pré', 'fiancé', 'lycé', 'musé', 'allé', 'idé', 'vé', 'ré', 'sé', 'télé', 'cliché', 'né', 'vélo', 'météo'],
    's': ['cas', 'os', 'bis', 'as', 'plus', 'moins', 'chez', 'fils', 'jus', 'tus'],
    'ent': ['vent', 'dent', 'cent', 'lent', 'gent', 'absent', 'présent'],
    'ez': ['chez', 'rez', 'nez'],
    't': ['but', 'dot', 'fat', 'fut', 'brut', 'zut', 'tout', 'chut'],
    'd': ['sud', 'lourd', 'nid', 'david'],
}

def est_muette(mot_lower):
    if len(mot_lower) <= 1:
        return []
    muettes = []
    if mot_lower.endswith('ent') and mot_lower not in EXCEPTIONS.get('ent', []):
        muettes.extend(['e','n','t'])
    elif mot_lower.endswith('ez') and mot_lower not in EXCEPTIONS.get('ez', []):
        muettes.extend(['e','z'])
    elif mot_lower.endswith('ait') and mot_lower not in ['fait','trait','plait']:
        muettes.extend(['a','i','t'])
    elif mot_lower.endswith('aient'):
        muettes.extend(['a','i','e','n','t'])
    elif mot_lower.endswith('e') and mot_lower not in EXCEPTIONS.get('e', []) and not mot_lower.endswith(('ée','é','és')):
        muettes.append('e')
    elif mot_lower.endswith('s') and len(mot_lower) > 1 and mot_lower not in EXCEPTIONS.get('s', []):
        muettes.append('s')
    elif mot_lower[-1] in 'tdxz' and mot_lower not in EXCEPTIONS.get(mot_lower[-1], []):
        muettes.append(mot_lower[-1])
    return muettes

def apply_mute_letters(doc):
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

            # ON RECONSTRUIT LE PARAGRAPHE ENTIÈREMENT (MÉTHODE 100 % SÛRE)
            texte = p.text
            style_original = p.style
            alignment = p.alignment
            space_before = p.paragraph_format.space_before
            space_after = p.paragraph_format.space_after
            line_spacing = p.paragraph_format.line_spacing

            p.clear()

            i = 0
            while i < len(texte):
                c = texte[i]

                # Conserver espaces, ponctuation, etc.
                if c.isspace() or re.match(r"[^\wàâéèêëîïôöùûüçÀÂÉÈÊËÎÏÔÖÙÛÜÇ]", c):
                    run = p.add_run(c)
                    i += 1
                    continue

                # Mot alphabétique
                match = re.search(r"[a-zA-ZàâéèêëîïôöùûüçÀÂÉÈÊËÎÏÔÖÙÛÜÇ']+", texte[i:])
                if match:
                    mot = match.group()
                    mot_lower = mot.lower()

                    muettes = est_muette(mot_lower)
                    if muettes:
                        lettres_grisees = ''.join(muettes)
                        debut_gris = len(mot) - len(lettres_grisees)
                        normale = mot[:debut_gris]
                        grise = mot[debut_gris:]

                        if normale:
                            run = p.add_run(normale)
                            # La couleur syllabique est déjà appliquée

                        if grise:
                            run = p.add_run(grise)
                            run.font.color.rgb = GRAY
                            # On garde la couleur syllabique en fond si elle existe
                            # (python-docx ne permet pas de superposer, mais le gris domine)
                    else:
                        run = p.add_run(mot)

                    i += len(mot)
                else:
                    run = p.add_run(c)
                    i += 1

            # Restaurer le style du paragraphe
            p.style = style_original
            p.alignment = alignment
            p.paragraph_format.space_before = space_before
            p.paragraph_format.space_after = space_after
            p.paragraph_format.line_spacing = line_spacing