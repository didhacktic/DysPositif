# core/syllables.py
import re
from docx.shared import RGBColor

COUL_SYLL = [RGBColor(220, 20, 60), RGBColor(30, 144, 255)]

WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ'’'-]+")

ACCENT_MAP = str.maketrans({
    'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
    'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
    'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
    'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y',
    'æ': 'ae', 'œ': 'oe',
    'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A', 'Å': 'A',
    'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
    'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
    'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O',
    'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U',
    'Ç': 'C', 'Ñ': 'N', 'Ý': 'Y', 'Ÿ': 'Y',
    'Æ': 'AE', 'Œ': 'OE'
})

def normalize(word: str) -> str:
    return word.lower().translate(ACCENT_MAP)

def split_syllables(word_norm: str):
    vowels = "aeiouy"
    parts = []
    current = ""
    for c in word_norm:
        current += c
        if c in vowels:
            parts.append(current)
            current = ""
    if current:
        if parts:
            parts[-1] += current
        else:
            parts.append(current)
    return parts or [word_norm]

def apply_syllables(doc):
    counter = [0]

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

            texte = p.text
            p.clear()
            i = 0

            while i < len(texte):
                c = texte[i]

                if c.isspace() or c in ".,;:!?()[]{}«»“”‘’'’/\\-–—*+=<>@#$%^&~":
                    p.add_run(c)
                    i += 1
                    continue

                match = WORD_PATTERN.match(texte[i:])
                if match:
                    mot = match.group()
                    mot_norm = normalize(mot)
                    parts = split_syllables(mot_norm)

                    pos = 0
                    for part in parts:
                        part_len = len(part)
                        orig_part = mot[pos:pos + part_len]
                        if normalize(orig_part) != part:
                            orig_part = mot[pos:pos + 1]
                            pos += 1
                        else:
                            pos += part_len

                        if orig_part:
                            run = p.add_run(orig_part)
                            run.font.color.rgb = COUL_SYLL[counter[0] % 2]
                            counter[0] += 1

                    i += len(mot)
                else:
                    p.add_run(c)
                    i += 1