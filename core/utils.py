# -------------------------------------------------
# core/utils.py – Outils partagés (police, styles, etc.)
# -------------------------------------------------
from docx.shared import Pt


def apply_font_consistently(doc, police_name: str, taille_pt_value: int):
    """
    Applique police + taille à TOUS les runs, même après p.clear()
    Fonctionne dans : corps, tableaux, zones de texte, en-têtes, pieds de page
    """
    taille_pt = Pt(taille_pt_value)

    def process_container(container):
        if container is None:
            return
        for p in container.paragraphs:
            # Appliquer à chaque run
            for run in p.runs:
                run.font.name = police_name
                run.font.size = taille_pt
            # Forcer aussi le style du paragraphe (évite les retours à 11pt)
            if p.style and hasattr(p.style, 'font'):
                try:
                    p.style.font.name = police_name
                    p.style.font.size = taille_pt
                except:
                    pass  # Certains styles sont verrouillés

    # Corps principal
    process_container(doc)

    # Tableaux
    for table in doc.tables:
        for cell in table._cells:
            process_container(cell)

    # Zones de texte (shapes)
    for shape in doc.inline_shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            process_container(shape.text_frame)

    # En-têtes et pieds de page
    for section in doc.sections:
        process_container(section.header)
        process_container(section.footer)
        if section.different_first_page_header_footer:
            process_container(section.first_page_header)
            process_container(section.first_page_footer)


def apply_spacing_and_line_spacing(doc, espacement: bool, interlignes: bool):
    """
    Applique espacement des lettres + interlignage 1.5
    """
    for p in doc.paragraphs:
        for run in p.runs:
            if espacement:
                run.font.spacing = Pt(2.4)
        if interlignes:
            p.paragraph_format.line_spacing = 1.5

    # Tableaux
    for table in doc.tables:
        for cell in table._cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if espacement:
                        run.font.spacing = Pt(2.4)
                if interlignes:
                    p.paragraph_format.line_spacing = 1.5

    # Zones de texte
    for shape in doc.inline_shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    if espacement:
                        run.font.spacing = Pt(2.4)
                if interlignes:
                    p.paragraph_format.line_spacing = 1.5