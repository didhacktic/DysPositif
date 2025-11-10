# -------------------------------------------------
# core/numbers_position.py – Couleur par position (u/d/c) – PRÉSERVE TOUT
# -------------------------------------------------
from docx.shared import RGBColor

COULEURS = [RGBColor(30,144,255), RGBColor(220,20,60), RGBColor(0,128,0)]

def apply_position_numbers(doc):
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
            runs_data = [(r.text, r.bold, r.italic, r.underline, r.font.color.rgb if r.font.color and r.font.color.rgb else None) for r in p.runs]
            p.clear()
            pos_digit = 0
            i = 0
            run_idx = 0
            while i < len(texte):
                c = texte[i]
                if c.isdigit():
                    r = p.add_run(c)
                    r.font.color.rgb = COULEURS[pos_digit % 3]
                    pos_digit += 1
                    if run_idx < len(runs_data) and runs_data[run_idx][0]:
                        r.bold, r.italic, r.underline = runs_data[run_idx][1:4]
                        if runs_data[run_idx][4]:
                            r.font.color.rgb = runs_data[run_idx][4]
                        runs_data[run_idx] = (runs_data[run_idx][0][1:],) + runs_data[run_idx][1:]
                        if not runs_data[run_idx][0]:
                            run_idx += 1
                else:
                    r = p.add_run(c)
                    if run_idx < len(runs_data) and runs_data[run_idx][0]:
                        r.bold, r.italic, r.underline = runs_data[run_idx][1:4]
                        if runs_data[run_idx][4]:
                            r.font.color.rgb = runs_data[run_idx][4]
                        runs_data[run_idx] = (runs_data[run_idx][0][1:],) + runs_data[run_idx][1:]
                        if not runs_data[run_idx][0]:
                            run_idx += 1
                i += 1