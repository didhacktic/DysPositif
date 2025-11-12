#!/usr/bin/env python
# -*- coding: UTF-8 -*-

"""
core/mute_letters.py
Grisage des lettres muettes — version robuste pour fonctionner correctement
après coloration syllabique (travaille au niveau caractère tout en conservant
les styles d'origine des runs).

Amélioration : meilleure détection des noms propres (spaCy ENT + fallback Titlecase)
pour éviter de griser des lettres finales sur des noms propres (ex. Jean-Louis).
"""

import re
import os
from typing import Set
from docx.shared import RGBColor

# Utiliser le même WORD_PATTERN que core.syllables pour une segmentation cohérente
from .syllables import WORD_PATTERN

# --- Couleur ---
GRIS = RGBColor(200, 200, 200)

# --- spaCy ---
try:
    import spacy

    # charge le modèle français si disponible
    nlp = spacy.load("fr_core_news_md")
    SPACY_OK = True
except Exception as e:
    # spaCy peut être absent dans certains environnements -> on garde un fallback heuristique
    print(f"spaCy erreur : {e}")
    SPACY_OK = False
    nlp = None

# --- regex léger pour "tous + déterminant/article"
ARTICLES = [
    "le", "la", "les", "un", "une", "des", "du", "de", "au", "aux",
    "mon", "ma", "mes", "ton", "ta", "tes", "son", "sa", "ses",
    "notre", "nos", "votre", "vos", "ce", "cet", "cette", "ces",
    "quelques", "chaque", "tout", "tous"
]
ARTICLES_RE = re.compile(r"\b[tT]ous\s+(?:" + "|".join(re.escape(a) for a in ARTICLES) + r")\b", re.IGNORECASE)


def is_tous_followed_by_article(sentence: str) -> bool:
    if not sentence:
        return False
    return bool(ARTICLES_RE.search(sentence))


# --- EXCEPTIONS (inchangées) ---
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

EXCEPTIONS_D = {"david"}  # David

CAS_PARTICULIERS = {
    "croc": "c", "crocs": "cs",
    "clef": "f", "clefs": "fs",
    "cerf": "f", "cerfs": "fs",
    "boeuf": "fs", "bœuf": "fs", "boeufs": "fs", "bœufs": "fs",
    "oeuf": "fs", "œuf": "fs", "oeufs": "fs", "œufs": "fs"
}


# --- HELPERS spaCy ---
def _prev_non_punct(doc, i, sent_start):
    j = i - 1
    while j >= sent_start:
        if not doc[j].is_punct:
            return doc[j]
        j -= 1
    return None


def _next_non_punct(doc, i, sent_end):
    j = i + 1
    while j < sent_end:
        if not doc[j].is_punct:
            return doc[j]
        j += 1
    return None


# --- FONCTIONS spaCy utilitaires (inchangées sauf is_proper_noun) ---
def is_verb(word: str, sentence: str) -> bool:
    if not SPACY_OK or not sentence:
        return False
    doc = nlp(sentence)
    for token in doc:
        if token.text.lower() == word.lower():
            return token.pos_ == "VERB"
    return False


def is_negation_plus(sentence: str, word: str) -> bool:
    if not SPACY_OK or not sentence:
        return False

    doc = nlp(sentence)
    idx_plus = None
    for i, t in enumerate(doc):
        if t.text.lower() == word.lower():
            idx_plus = i
            break
    if idx_plus is None:
        return False

    plus_tok = doc[idx_plus]
    sent = plus_tok.sent
    sent_start = sent.start
    sent_end = sent.end

    max_prev = 4
    prev_tokens = []
    j = plus_tok.i - 1
    while j >= sent_start and len(prev_tokens) < max_prev:
        if not doc[j].is_punct:
            prev_tokens.append(doc[j])
        j -= 1

    neg_tokens = [tok for tok in prev_tokens if tok.text.lower() in {"ne", "n'"} or getattr(tok, "dep_", "") == "neg"]
    if neg_tokens:
        neg_tok = neg_tokens[0]
        if any(tok.pos_ == "VERB" for tok in sent if neg_tok.i < tok.i < plus_tok.i):
            return True
        next_tok = _next_non_punct(doc, plus_tok.i, sent_end)
        if next_tok is not None and next_tok.pos_ == "VERB":
            return True

    return False


def is_plus_relevant(sentence: str, word: str) -> bool:
    if not SPACY_OK or not sentence:
        return False

    doc = nlp(sentence)
    for t in doc:
        if t.text.lower() == word.lower():
            sent = t.sent
            left = _prev_non_punct(doc, t.i, sent.start)
            right = _next_non_punct(doc, t.i, sent.end)

            def is_number(tok):
                return getattr(tok, "like_num", False) or tok.pos_ == "NUM"

            if left is not None and right is not None:
                if is_number(left) and is_number(right):
                    return False
                if left.pos_ == "PRON" and right.pos_ == "PRON":
                    return False
                if left.pos_ in {"NOUN", "PROPN"} and right.pos_ in {"NOUN", "PROPN"}:
                    return False

            return is_negation_plus(sentence, word)
    return False


def is_tous_determiner(sentence: str, word: str) -> bool:
    return is_tous_followed_by_article(sentence)


def is_proper_noun(sentence: str, word: str) -> bool:
    """
    Détection renforcée d'un nom propre pour `word` dans `sentence`.
    - Si spaCy est disponible, on considère PROPN ou entités PERSON comme nom propre.
    - En fallback (spaCy absent ou non concluant) on utilise heuristique Titlecase
      (première lettre majuscule) pour éviter de griser les noms propres après
      syllabation (ex. Louis dans Jean-Louis).
    """
    if not sentence:
        return False

    # Première tentative : spaCy si disponible
    if SPACY_OK:
        try:
            doc = nlp(sentence)
            for token in doc:
                if token.text.lower() == word.lower():
                    # token.pos_ == 'PROPN' -> nom propre
                    if token.pos_ == "PROPN":
                        return True
                    # si spaCy a identifié une entité PERSON, aussi valide
                    if getattr(token, "ent_type_", "") in {"PER", "PERSON"}:
                        return True
            # pas trouvé comme PROPN/entité
        except Exception:
            # si spaCy plante, on passera au fallback heuristique
            pass

    # Fallback heuristique : mot commençant par majuscule (Titlecase)
    # Note : peut donner de faux positifs pour début de phrase mais c'est un compromis
    # acceptable si spaCy n'est pas fiable sur ce token.
    if word and word[0].isupper():
        return True

    return False


def get_mute_positions(word: str, sentence: str = None) -> Set[int]:
    w = word.lower()
    positions = set()

    # Ne jamais griser les noms propres (si détecté)
    if sentence and is_proper_noun(sentence, word):
        return positions

    # Cas particuliers
    if w in CAS_PARTICULIERS:
        for c in CAS_PARTICULIERS[w]:
            idx = w.rfind(c)
            if idx != -1:
                positions.add(idx)
        return positions

    # Règle : h début
    if w and w[0] == "h":
        positions.add(0)

    # ent + verbe
    if w.endswith("ent") and sentence and is_verb(w, sentence):
        positions.add(len(w) - 2)
        positions.add(len(w) - 1)

    # plus + négation
    if w == "plus" and sentence and is_plus_relevant(sentence, w):
        positions.add(len(w) - 1)
        return positions

    # tous + déterminant
    if w == "tous" and sentence and is_tous_followed_by_article(sentence):
        positions.add(len(w) - 1)
        return positions

    # aient -> ent
    if w.endswith("aient") and w != "aient":
        positions.add(len(w) - 3)
        positions.add(len(w) - 2)
        positions.add(len(w) - 1)
        return positions

    # Règles finales
    last = len(w) - 1
    if last < 0:
        return positions

    if w[last] == "d" and w in EXCEPTIONS_D:
        pass
    elif w[last] == "d":
        positions.add(last)

    if w[last] == "b" and w not in EXCEPTIONS_B:
        positions.add(last)

    # terminaisons 'ie','ée' et 'ue' (avec exclusion 'gue'/'que')
    if w.endswith(("ie", "ée")):
        positions.add(last)
    elif w.endswith("ue"):
        if not (w.endswith("gue") or w.endswith("que")):
            positions.add(last)

    if w[last] == "g" and w not in EXCEPTIONS_G:
        positions.add(last)
    if w[last] == "p" and w not in EXCEPTIONS_P:
        positions.add(last)
    if w[last] == "t" and w not in EXCEPTIONS_T:
        positions.add(last)
    if w[last] == "x" and w not in EXCEPTIONS_X:
        positions.add(last)
    if w[last] == "s" and w not in EXCEPTIONS_S:
        positions.add(last)
        if len(w) > 1:
            prev = w[:-1]
            prev_pos = get_mute_positions(prev, sentence)
            for p in prev_pos:
                positions.add(p)

    return positions


def copy_style(src, dst):
    for attr in ["bold", "italic", "underline"]:
        value = getattr(src, attr, None)
        if value is not None:
            try:
                setattr(dst, attr, value)
            except Exception:
                pass
    try:
        if src.font.name:
            dst.font.name = src.font.name
    except Exception:
        pass
    try:
        if src.font.size:
            dst.font.size = src.font.size
    except Exception:
        pass
    try:
        if getattr(src.font, "color", None) and getattr(src.font.color, "rgb", None):
            dst.font.color.rgb = src.font.color.rgb
    except Exception:
        pass


def apply_mute_letters(doc):
    """
    Version robuste : travaille sur le texte complet de chaque paragraphe en
    conservant pour chaque caractère le style d'origine. Calcule un mask des
    caractères à griser puis reconstruit des runs groupés par style+muted.

    Protection : si spaCy détecte PROPN '-' PROPN on désactive le mask pour la plage.
    """
    counter = 0

    HYPHENS = {"-", "–", "—", "\u2010", "\u2011"}

    for paragraph in doc.paragraphs:
        runs = list(paragraph.runs)
        if not runs:
            continue

        # Construire liste des caractères avec style associé
        chars = []
        for run in runs:
            text = run.text or ""
            style = {
                "bold": bool(run.bold),
                "italic": bool(run.italic),
                "underline": bool(run.underline),
                "font_name": getattr(run.font, "name", None),
                "font_size": getattr(run.font, "size", None),
                "color": None,
            }
            try:
                if getattr(run.font, "color", None) is not None:
                    rgb = getattr(run.font.color, "rgb", None)
                    if rgb is not None:
                        style["color"] = rgb
            except Exception:
                style["color"] = None

            for ch in text:
                chars.append((ch, style))

        if not chars:
            continue

        full_text = "".join(ch for ch, _ in chars)

        # mask booleen des caractères à griser
        mute_mask = [False] * len(full_text)

        # itérer sur les mots selon WORD_PATTERN
        for m in WORD_PATTERN.finditer(full_text):
            start, end = m.start(), m.end()
            word = full_text[start:end]
            try:
                mute_pos = get_mute_positions(word, full_text)
            except Exception:
                mute_pos = set()
            for p in mute_pos:
                idx = start + p
                if 0 <= idx < len(full_text):
                    mute_mask[idx] = True

        # Protection PROPN '-' PROPN : désactiver mask sur intervals détectés
        if SPACY_OK and full_text.strip():
            try:
                doc_sp = nlp(full_text)
                for i in range(len(doc_sp) - 2):
                    t0, t1, t2 = doc_sp[i], doc_sp[i + 1], doc_sp[i + 2]
                    if t0.pos_ == "PROPN" and t2.pos_ == "PROPN" and t1.text in HYPHENS:
                        span_start = t0.idx
                        span_end = t2.idx + len(t2.text)
                        for k in range(span_start, span_end):
                            if 0 <= k < len(mute_mask):
                                mute_mask[k] = False
            except Exception:
                pass

        # Reconstruire le paragraphe
        paragraph.clear()
        i = 0
        while i < len(full_text):
            ch, style = chars[i]
            muted = mute_mask[i]

            j = i + 1
            while j < len(full_text):
                if chars[j][1] != style or mute_mask[j] != muted:
                    break
                j += 1

            segment = "".join(chars[k][0] for k in range(i, j))
            new_run = paragraph.add_run(segment)

            try:
                if style["font_name"]:
                    new_run.font.name = style["font_name"]
            except Exception:
                pass
            try:
                if style["font_size"]:
                    new_run.font.size = style["font_size"]
            except Exception:
                pass
            try:
                new_run.bold = style["bold"]
                new_run.italic = style["italic"]
                new_run.underline = style["underline"]
            except Exception:
                pass

            try:
                if muted:
                    new_run.font.color.rgb = GRIS
                    counter += (j - i)
                else:
                    if style["color"] is not None:
                        new_run.font.color.rgb = style["color"]
            except Exception:
                pass

            i = j

    return counter