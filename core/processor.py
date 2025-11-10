# -------------------------------------------------
# core/processor.py – Point d’entrée unique (compatible avec main.py)
# -------------------------------------------------
from .formatter import format_document

# Ancienne signature attendue par main.py : process_document(filepath, progress_callback)
# Nouvelle signature : format_document(filepath) → update_progress est global

def process_document(filepath: str, progress_callback=None):
    # On ignore progress_callback car update_progress est maintenant global
    format_document(filepath)