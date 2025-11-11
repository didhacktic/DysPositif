#!/usr/bin/env python
# -*- coding: UTF-8 -*-

"""
mute_letters.py – Grisage des lettres muettes
Règles utilisateur implémentées, sans découpage syllabique
NOUVELLE RÈGLE : aient → griser ent (sauf si mot = "aient")
"""

import re
import os
from docx.shared import RGBColor

# --- Couleur ---
GRIS = RGBColor(200, 200, 200)

# --- spaCy ---
try:
    import spacy
    nlp = spacy.load("fr_core_news_md")  # ← MODÈLE MOYEN
    SPACY_OK = True
except Exception as e:
    print(f"spaCy erreur : {e}")
    SPACY_OK = False
    nlp = None

# --- EXCEPTIONS ---
EXCEPTIONS_B = {
    "rib", "blob", "club", "pub", "kebab", "nabab", "snob", "toubib",
    "baobab", "jazzclub", "motoclub", "night-club"
}
EXCEPTIONS_G = {
    "grog", "ring", "bang", "gong", "yang", "ying", "slang", "gang", "erg",
    "iceberg", "zig", "zigzag", "krieg", "bowling", "briefing", "shopping",
    "building", "camping", "parking", "living", "marketing", "dancing",
    "jogging", "surfing", "training", "meeting", "feeling", "holding",
    "standing", "trading"
}
EXCEPTIONS_P = {
    "stop", "workshop", "handicap", "wrap", "ketchup", "top", "flip-flop",
    "hip-hop", "clip", "slip", "trip", "grip", "strip", "shop", "drop",
    "hop", "pop", "flop", "chop", "prop", "crop", "laptop", "desktop"
}
EXCEPTIONS_T = {
    "but", "chut", "fiat", "brut", "concept", "foot", "huit", "mat", "net",
    "ouest", "rut", "out", "ut", "flirt", "kurt", "loft", "raft", "rift",
    "soft", "watt", "west", "abstract", "affect", "apart", "audit", "belt",
    "best", "blast", "boost", "compact", "connect", "contact", "correct",
    "cost", "craft", "cut", "direct", "district", "draft", "drift", "exact",
    "exit", "impact", "infect", "input", "must", "next", "night", "outfit",
    "output", "paint", "perfect", "plot", "post", "print", "prompt",
    "prospect", "react", "root", "set", "shirt", "short", "shot", "smart",
    "spirit", "split", "spot", "sprint", "start", "strict", "tact", "test",
    "tilt", "tract", "trust", "twist", "volt", "et", "est"
}
EXCEPTIONS_X = {
    "six", "dix", "index", "duplex", "latex", "lynx", "matrix", "mix",
    "multiplex", "reflex", "relax", "remix", "silex", "thorax", "vortex", "xerox"
}
EXCEPTIONS_S = {
    "bus", "ours", "tous", "plus", "ars", "cursus", "lapsus", "virus",
    "cactus", "consensus", "us", "as", "mas", "bis", "lys", "métis", "os",
    "bonus", "campus", "focus", "boss", "stress", "express", "dress",
    "fitness", "Arras", "s", "houmous", "humus", "humérus", "cubitus", "habitus", 
    "hiatus", "des", "mes", "tes", "ces", "les", "ses"  
}

# --- EXCEPTIONS SUPPLÉMENTAIRES ---
EXCEPTIONS_D = {"david"}  # David

# --- CAS PARTICULIERS ---
CAS_PARTICULIERS = {
    "croc": "c", "crocs": "cs",
    "clef": "f", "clefs": "fs",
    "cerf": "f", "cerfs": "fs",
    "boeuf": "fs", "bœuf": "fs", "boeufs": "fs", "bœufs": "fs",
    "oeuf": "fs", "œuf": "fs", "oeufs": "fs", "œufs": "fs"
}

# --- FONCTIONS ---
def is_verb(word, sentence):
    if not SPACY_OK:
        return False
    doc = nlp(sentence)
    for token in doc:
        if token.text.lower() == word.lower():
            return token.pos_ == "VERB"
    return False

def is_negation_plus(sentence, word):
    if not SPACY_OK:
        return False
    doc = nlp(sentence)
    tokens = [t.text.lower() for t in doc]
    try:
        idx = tokens.index(word.lower())
        return any(t in {"ne", "n'"} for t in tokens[max(0, idx-5):idx])
    except ValueError:
        return False

def get_mute_positions(word, sentence=None):
    w = word.lower()
    positions = set()

    # Cas particuliers
    if w in CAS_PARTICULIERS:
        for c in CAS_PARTICULIERS[w]:
            idx = w.rfind(c)
            if idx != -1:
                positions.add(idx)
        return positions

    # Règle 1: h début
    if w and w[0] == 'h':
        positions.add(0)

    # Règle 12: ent + verbe
    if w.endswith('ent') and sentence and is_verb(w, sentence):
        positions.add(len(w)-2)  # n
        positions.add(len(w)-1)  # t
        #return positions

    # Règle 13: plus + négation
    if w == "plus" and sentence and is_negation_plus(sentence, w):
        positions.add(len(w)-1)
        return positions

    # NOUVELLE RÈGLE : aient → griser ent (sauf si mot = "aient")
    if w.endswith('aient') and w != "aient":
        positions.add(len(w)-3)  # e
        positions.add(len(w)-2)  # n
        positions.add(len(w)-1)  # t
        return positions

    # Règles finales
    last = len(w) - 1
    if last < 0:
        return positions

    # Exception d
    if w[last] == 'd' and w in EXCEPTIONS_D:
        pass
    elif w[last] == 'd':
        positions.add(last)

    if w[last] == 'b' and w not in EXCEPTIONS_B:
        positions.add(last)
    if w.endswith(('ie', 'ée')):
        positions.add(last)
    if w[last] == 'g' and w not in EXCEPTIONS_G:
        positions.add(last)
    if w[last] == 'p' and w not in EXCEPTIONS_P:
        positions.add(last)
    if w[last] == 't' and w not in EXCEPTIONS_T:
        positions.add(last)
    if w[last] == 'x' and w not in EXCEPTIONS_X:
        positions.add(last)
    if w[last] == 's' and w not in EXCEPTIONS_S:
        positions.add(last)
        # Règle 11: lettre précédente
        if len(w) > 1:
            prev = w[:-1]
            prev_pos = get_mute_positions(prev, sentence)
            for p in prev_pos:
                positions.add(p)

    return positions

def copy_style(src, dst):
    for attr in ['bold', 'italic', 'underline']:
        value = getattr(src, attr, None)
        if value is not None:
            setattr(dst, attr, value)
    if src.font.name:
        dst.font.name = src.font.name
    if src.font.size:
        dst.font.size = src.font.size
    if src.font.color.rgb:
        dst.font.color.rgb = src.font.color.rgb

def apply_mute_letters(doc):
    counter = 0
    for paragraph in doc.paragraphs:
        sentence = paragraph.text
        runs = list(paragraph.runs)
        paragraph.clear()  # Vider sans perdre le paragraphe

        for run in runs:
            text = run.text
            if not text.strip():
                new_run = paragraph.add_run(text)
                copy_style(run, new_run)
                continue

            pos = 0
            word_regex = re.compile(r"\b\w+\b")
            for match in word_regex.finditer(text):
                start, end = match.start(), match.end()
                word = text[start:end]
                mute_pos = get_mute_positions(word, sentence)

                # Préfixe
                if start > pos:
                    prefix = text[pos:start]
                    new_run = paragraph.add_run(prefix)
                    copy_style(run, new_run)
                pos = start

                # Mot
                local_pos = 0
                for idx in sorted(mute_pos):
                    rel_start = idx - local_pos
                    if rel_start > 0:
                        part = word[local_pos:idx]
                        new_run = paragraph.add_run(part)
                        copy_style(run, new_run)
                    mute_char = word[idx]
                    mute_run = paragraph.add_run(mute_char)
                    copy_style(run, mute_run)
                    mute_run.font.color.rgb = GRIS
                    counter += 1
                    local_pos = idx + 1

                if local_pos < len(word):
                    rest = word[local_pos:]
                    new_run = paragraph.add_run(rest)
                    copy_style(run, new_run)

                pos = end

            # Suffixe
            if pos < len(text):
                suffix = text[pos:]
                new_run = paragraph.add_run(suffix)
                copy_style(run, new_run)

    return counter