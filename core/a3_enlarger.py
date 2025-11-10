# -------------------------------------------------
# core/a3_enlarger.py – Format A3 + agrandissement objets
# -------------------------------------------------
from docx.shared import Mm

def apply_a3_format(doc, agrandir_objets):
    s = doc.sections[0]
    s.page_height = Mm(420)
    s.page_width = Mm(297)
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Mm(20)

    if agrandir_objets:
        fw, fh = 1.40, 1.25
        for table in doc.tables:
            for col in table.columns:
                if col.width:
                    col.width = Mm(col.width.mm * fw)
            for row in table.rows:
                if row.height:
                    row.height = Mm(row.height.mm * fh)
        for shape in doc.inline_shapes:
            if hasattr(shape, 'width') and shape.width:
                shape.width = int(shape.width * fw)
                shape.height = int(shape.height * fh)