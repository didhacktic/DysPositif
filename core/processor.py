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


def _convert_vml_to_drawingml(docx_path: str):
    """
    Conversion des zones de texte VML en DrawingML moderne (toutes les occurrences).
    Deux passes : collecte des cibles puis remplacement pour éviter les effets d'itération.
    """
    import zipfile, tempfile, shutil, re
    from lxml import etree

    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            zin.extractall(temp_dir)
        doc_xml = os.path.join(temp_dir, 'word', 'document.xml')
        tree = etree.parse(doc_xml)
        root = tree.getroot()

        # Namespaces (utiliser URI sans accolades pour création)
        v_uri = 'urn:schemas-microsoft-com:vml'
        w_uri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        wp_uri = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        a_uri = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        wps_uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

        targets = []  # (parent, index, style, txbx_content)

        # Collecte
        for parent in root.iter():
            children = list(parent)
            for idx, child in enumerate(children):
                if child.tag != f'{{{w_uri}}}pict':
                    continue
                vml_shape = child.find(f'.//{{{v_uri}}}shape')
                if vml_shape is None:
                    continue
                # Chercher v:textbox + w:txbxContent
                v_textbox = vml_shape.find(f'./{{{v_uri}}}textbox')
                if v_textbox is None:
                    continue
                txbx_content = v_textbox.find(f'{{{w_uri}}}txbxContent')
                if txbx_content is None:
                    continue
                style = vml_shape.get('style', '')
                if not all(k in style for k in ('margin-left','margin-top','width','height')):
                    continue
                targets.append((parent, idx, style, txbx_content))

        if not targets:
            print('  Aucune textbox VML trouvée')
            return

        modifications = 0
        shape_id = 1
        for parent, idx, style, txbx_content in targets:
            ml = re.search(r'margin-left:([0-9.]+)pt', style)
            mt = re.search(r'margin-top:([0-9.]+)pt', style)
            w_match = re.search(r'width:([0-9.]+)pt', style)
            h_match = re.search(r'height:([0-9.]+)pt', style)
            if not (ml and mt and w_match and h_match):
                continue
            ml_pt = float(ml.group(1)); mt_pt = float(mt.group(1))
            w_pt = float(w_match.group(1)); h_pt_orig = float(h_match.group(1))
            h_pt = h_pt_orig * 5.0
            # EMUs
            factor = 914400/72.0
            ml_emu = int(ml_pt * factor); mt_emu = int(mt_pt * factor)
            w_emu = int(w_pt * factor); h_emu = int(h_pt * factor)

            drawing = etree.Element(f'{{{w_uri}}}drawing')
            anchor = etree.SubElement(drawing, f'{{{wp_uri}}}anchor', behindDoc='1', distT='0', distB='0', distL='0', distR='0',
                                       simplePos='0', locked='0', layoutInCell='0', allowOverlap='1', relativeHeight='5')
            etree.SubElement(anchor, f'{{{wp_uri}}}simplePos', x='0', y='0')
            pos_h = etree.SubElement(anchor, f'{{{wp_uri}}}positionH', relativeFrom='page')
            etree.SubElement(pos_h, f'{{{wp_uri}}}posOffset').text = str(ml_emu)
            pos_v = etree.SubElement(anchor, f'{{{wp_uri}}}positionV', relativeFrom='paragraph')
            etree.SubElement(pos_v, f'{{{wp_uri}}}posOffset').text = str(mt_emu)
            etree.SubElement(anchor, f'{{{wp_uri}}}extent', cx=str(w_emu), cy=str(h_emu))
            etree.SubElement(anchor, f'{{{wp_uri}}}effectExtent', l='3810', t='3810', r='2540', b='2540')
            etree.SubElement(anchor, f'{{{wp_uri}}}wrapNone')
            etree.SubElement(anchor, f'{{{wp_uri}}}docPr', id=str(shape_id), name=f'Textbox {shape_id}')

            graphic = etree.SubElement(anchor, f'{{{a_uri}}}graphic')
            gdata = etree.SubElement(graphic, f'{{{a_uri}}}graphicData', uri='http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
            wsp = etree.SubElement(gdata, f'{{{wps_uri}}}wsp')
            etree.SubElement(wsp, f'{{{wps_uri}}}cNvSpPr')
            sp_pr = etree.SubElement(wsp, f'{{{wps_uri}}}spPr')
            xfrm = etree.SubElement(sp_pr, f'{{{a_uri}}}xfrm')
            etree.SubElement(xfrm, f'{{{a_uri}}}off', x='0', y='0')
            etree.SubElement(xfrm, f'{{{a_uri}}}ext', cx=str(w_emu), cy=str(h_emu))
            geom = etree.SubElement(sp_pr, f'{{{a_uri}}}prstGeom', prst='rect')
            etree.SubElement(geom, f'{{{a_uri}}}avLst')
            etree.SubElement(sp_pr, f'{{{a_uri}}}noFill')
            ln = etree.SubElement(sp_pr, f'{{{a_uri}}}ln', w='6480')
            sf = etree.SubElement(ln, f'{{{a_uri}}}solidFill')
            etree.SubElement(sf, f'{{{a_uri}}}srgbClr', val='000000')
            etree.SubElement(ln, f'{{{a_uri}}}round')
            txbx = etree.SubElement(wsp, f'{{{wps_uri}}}txbx')
            txbx.append(txbx_content)
            body_pr = etree.SubElement(wsp, f'{{{wps_uri}}}bodyPr', lIns='0', rIns='0', tIns='0', bIns='0', anchor='t')
            etree.SubElement(body_pr, f'{{{a_uri}}}noAutofit')

            # Remplacement
            parent[idx] = drawing
            modifications += 1
            print(f"  Textbox {modifications} convertie: {w_pt:.1f}×{h_pt_orig:.1f}pt VML → {w_pt:.1f}×{h_pt:.1f}pt DrawingML")
            shape_id += 1

        if modifications:
            print(f"  → {modifications} textbox(es) convertie(s) VML→DrawingML")
            tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
            with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for rdir, dirs, files in os.walk(temp_dir):
                    for f in files:
                        p = os.path.join(rdir, f)
                        zout.write(p, os.path.relpath(p, temp_dir))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


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
    
    # Post-traitement: conversion VML → DrawingML pour textboxes
    try:
        _convert_vml_to_drawingml(output)
    except Exception as e:
        print(f"Erreur conversion zones de texte: {e}")
    
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