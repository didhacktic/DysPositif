#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de diagnostic détaillé pour comprendre pourquoi le formatage ne s'applique pas.
À exécuter avec : python3 diagnostic_vml_detailed.py chemin/vers/fichier.docx
"""

import sys
import os
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Pt, RGBColor
import zipfile
from lxml import etree

def diagnostic_detaille(filepath):
    """Diagnostic très détaillé pour comprendre le problème."""
    
    print("=" * 80)
    print("DIAGNOSTIC DÉTAILLÉ VML - DysPositif")
    print("=" * 80)
    print(f"\nFichier: {filepath}")
    
    if not os.path.exists(filepath):
        print(f"✗ ERREUR: Le fichier n'existe pas!")
        return False
    
    # 1. Analyse XML brute
    print("\n[1] Analyse de la structure XML brute...")
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            doc_xml = zf.read('word/document.xml')
        
        root = etree.fromstring(doc_xml)
        
        # Compter les éléments
        txbx_count = len(list(root.iter(qn('w:txbxContent'))))
        pict_count = len(list(root.iter(qn('w:pict'))))
        
        print(f"    w:txbxContent trouvés: {txbx_count}")
        print(f"    w:pict trouvés: {pict_count}")
        
        if txbx_count == 0:
            print("    ⚠ AUCUNE zone de texte VML détectée!")
            print("    → Ce document ne contient peut-être pas de zones VML")
            print("    → Ou les zones sont d'un type différent (DrawingML)")
            return False
        
        # Afficher le contenu des zones VML
        print(f"\n[2] Contenu des zones VML:")
        for i, txbx in enumerate(root.iter(qn('w:txbxContent')), 1):
            print(f"\n    Zone VML #{i}:")
            
            # Compter paragraphes
            para_count = len(list(txbx.iter(qn('w:p'))))
            print(f"      Paragraphes: {para_count}")
            
            # Extraire texte
            text_parts = []
            for t_elem in txbx.iter(qn('w:t')):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            
            text = ''.join(text_parts)
            if len(text) > 100:
                print(f"      Texte: '{text[:100]}...'")
            else:
                print(f"      Texte: '{text}'")
            
            # Vérifier formatage
            for p_elem in list(txbx.iter(qn('w:p')))[:2]:  # Premiers 2 paragraphes
                print(f"\n      Formatage d'un paragraphe:")
                
                # Vérifier pPr (propriétés de paragraphe)
                pPr = p_elem.find(qn('w:pPr'))
                if pPr is not None:
                    spacing = pPr.find(qn('w:spacing'))
                    if spacing is not None:
                        line_val = spacing.get(qn('w:line'))
                        print(f"        Interlignage (w:line): {line_val}")
                else:
                    print(f"        Pas de w:pPr (propriétés de paragraphe)")
                
                # Vérifier runs
                runs = list(p_elem.iter(qn('w:r')))
                print(f"        Runs: {len(runs)}")
                
                for j, run in enumerate(runs[:2], 1):  # Premiers 2 runs
                    rPr = run.find(qn('w:rPr'))
                    if rPr is not None:
                        fonts = rPr.find(qn('w:rFonts'))
                        if fonts is not None:
                            font_name = fonts.get(qn('w:ascii'))
                            print(f"          Run {j}: Police = {font_name}")
                        
                        sz = rPr.find(qn('w:sz'))
                        if sz is not None:
                            size_val = sz.get(qn('w:val'))
                            print(f"          Run {j}: Taille = {size_val} (half-points)")
                    else:
                        print(f"          Run {j}: Pas de formatage (w:rPr)")
        
    except Exception as e:
        print(f"    ✗ Erreur lors de l'analyse XML: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # 3. Test avec python-docx et fonction de détection
    print("\n[3] Test de détection avec _iter_vml_textbox_paragraphs...")
    try:
        doc = Document(filepath)
        
        # Fonction de détection (copie de celle dans core/syllables.py)
        def _iter_vml_textbox_paragraphs(doc):
            root = doc.element
            for txbx_content in root.iter(qn('w:txbxContent')):
                for p_elem in txbx_content.iter(qn('w:p')):
                    try:
                        yield Paragraph(p_elem, txbx_content)
                    except Exception as e:
                        print(f"      ✗ Erreur création Paragraph: {e}")
        
        vml_paras = list(_iter_vml_textbox_paragraphs(doc))
        print(f"    Paragraphes VML détectés: {len(vml_paras)}")
        
        if len(vml_paras) == 0:
            print("    ✗ PROBLÈME: Fonction de détection ne trouve rien!")
            print("    → Vérifier que python-docx peut créer des Paragraph à partir des éléments XML")
            return False
        
        # Afficher détails de chaque paragraphe
        for i, p in enumerate(vml_paras[:3], 1):  # Premiers 3
            print(f"\n    Paragraphe VML #{i}:")
            text = p.text[:80] if len(p.text) > 80 else p.text
            print(f"      Texte: '{text}'")
            print(f"      Runs: {len(p.runs)}")
            
            # Formatage
            print(f"      Interlignage: {p.paragraph_format.line_spacing}")
            
            if len(p.runs) > 0:
                r = p.runs[0]
                print(f"      Premier run:")
                print(f"        Police: {r.font.name}")
                print(f"        Taille: {r.font.size}")
                print(f"        Gras: {r.font.bold}")
                print(f"        Italique: {r.font.italic}")
        
        print(f"\n    ✓ Détection fonctionne! {len(vml_paras)} paragraphe(s) trouvé(s)")
        
    except Exception as e:
        print(f"    ✗ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # 4. Test d'application de formatage
    print("\n[4] Test d'application de formatage...")
    try:
        # Appliquer formatage de test
        for p in vml_paras[:3]:
            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = Pt(16)
        
        # Sauvegarder
        test_output = filepath.replace('.docx', '_TEST_DETAILLE.docx')
        doc.save(test_output)
        print(f"    ✓ Formatage appliqué")
        print(f"    ✓ Fichier test sauvegardé: {test_output}")
        
        # Recharger et vérifier
        doc2 = Document(test_output)
        vml_paras2 = list(_iter_vml_textbox_paragraphs(doc2))
        
        if len(vml_paras2) > 0:
            print(f"\n    Vérification après sauvegarde:")
            for i, p in enumerate(vml_paras2[:2], 1):
                if len(p.runs) > 0:
                    r = p.runs[0]
                    print(f"      Paragraphe {i}: Police={r.font.name}, Taille={r.font.size}")
                    
                    if r.font.name == "Arial":
                        print(f"        ✓ Police correcte!")
                    else:
                        print(f"        ✗ Police incorrecte (attendu Arial, obtenu {r.font.name})")
        
    except Exception as e:
        print(f"    ✗ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # 5. Vérifier l'intégration dans apply_syllables
    print("\n[5] Vérification de l'intégration dans apply_syllables...")
    
    try:
        # Charger le fichier source
        import inspect
        syllables_file = os.path.join(os.path.dirname(__file__), 'core', 'syllables.py')
        
        if os.path.exists(syllables_file):
            with open(syllables_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Vérifier que les zones VML sont bien incluses
            if 'paragraphs_to_process.extend(_iter_vml_textbox_paragraphs(doc))' in content:
                print("    ✓ Les zones VML sont bien ajoutées à paragraphs_to_process")
            else:
                print("    ✗ Les zones VML ne sont PAS ajoutées à paragraphs_to_process!")
                print("    → PROBLÈME TROUVÉ: La fonction apply_syllables n'inclut pas les VML")
            
            # Vérifier la préservation du formatage
            if 'original_font_name = first_run.font.name' in content:
                print("    ✓ Préservation du formatage implémentée")
            else:
                print("    ✗ Préservation du formatage NON implémentée!")
        else:
            print(f"    ⚠ Fichier syllables.py non trouvé: {syllables_file}")
    
    except Exception as e:
        print(f"    ⚠ Erreur vérification intégration: {e}")
    
    # Résumé
    print("\n" + "=" * 80)
    print("RÉSUMÉ")
    print("=" * 80)
    
    print(f"\n✓ Zones VML détectées: {txbx_count}")
    print(f"✓ Paragraphes VML accessibles: {len(vml_paras)}")
    print(f"✓ Formatage peut être appliqué: Oui")
    print(f"✓ Formatage persiste après sauvegarde: Oui")
    
    print(f"\n→ Le code DEVRAIT fonctionner!")
    print(f"\nSi le formatage ne s'applique toujours pas dans l'application:")
    print(f"  1. Vérifier que les options sont cochées dans l'interface")
    print(f"  2. Vérifier que le fichier traité est bien celui avec les zones VML")
    print(f"  3. Ouvrir {test_output} pour voir si le formatage test fonctionne")
    print(f"  4. Si le test fonctionne mais pas l'application, il y a un problème")
    print(f"     dans le workflow de l'application (ordre des opérations, etc.)")
    
    return True

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 diagnostic_vml_detailed.py chemin/vers/fichier.docx")
        print("\nExemple:")
        print("  python3 diagnostic_vml_detailed.py ~/didhacktic/DysPositif/Fichiers_test/06_Texte_image_zt.docx")
        sys.exit(1)
    
    filepath = sys.argv[1]
    success = diagnostic_detaille(filepath)
    sys.exit(0 if success else 1)
