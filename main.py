# --------------------------------------------
# main.py – Dys’Positif – Interface principale 
# --------------------------------------------
import os
import sys
import threading
from tkinter import Tk, filedialog, messagebox

# Imports locaux
from ui.interface import create_interface
from core.processor import process_document
from converters.pdf_to_docx import pdf_to_docx
from converters.odt_to_docx import odt_to_docx
from utils.adobe_check import check_adobe_status


# -------------------------------------------------
# Configuration fenêtre principale
# -------------------------------------------------
root = Tk()
root.title("Dys’Positif – Adaptation pour la dyslexie")
root.geometry("900x700")
root.configure(bg="#f8f9fa")
root.resizable(False, False)

# Variable globale
current_file = None


# -------------------------------------------------
# Mise à jour barre de progression (globale via ui.interface)
# -------------------------------------------------
# → update_progress est importé depuis ui.interface (déjà global)
from ui.interface import update_progress


# -------------------------------------------------
# Fonction de continuation après conversion
# -------------------------------------------------
def continue_processing(docx_path: str):
    update_progress(30, "Conversion terminée. Application des adaptations dyslexie...")
    threading.Thread(
        target=process_document,
        args=(docx_path,),
        daemon=True
    ).start()


# -------------------------------------------------
# Traitement du fichier sélectionné (THREADS POUR CONVERSIONS)
# -------------------------------------------------
def handle_file(filepath: str):
    global current_file
    current_file = filepath
    ext = os.path.splitext(filepath)[1].lower()

    update_progress(5, "Analyse du fichier...")

    # --- PDF → DOCX via Adobe (THREAD SÉPARÉ) ---
    if ext == ".pdf":
        update_progress(10, "Conversion PDF → DOCX en cours (Adobe)...")
        def thread_pdf():
            docx_path = pdf_to_docx(filepath)
            if not docx_path:
                root.after(0, lambda: update_progress(0, "Échec conversion PDF"))
                root.after(0, lambda: messagebox.showerror(
                    "Erreur PDF",
                    "Impossible de convertir le PDF.\n"
                    "• Vérifiez vos identifiants Adobe\n"
                    "• Vérifiez votre connexion internet\n"
                    "• Essayez avec un PDF plus simple"
                ))
                return
            root.after(0, lambda: continue_processing(docx_path))
        threading.Thread(target=thread_pdf, daemon=True).start()
        return

    # --- ODT → DOCX via LibreOffice (THREAD SÉPARÉ) ---
    elif ext == ".odt":
        update_progress(10, "Conversion ODT → DOCX en cours (LibreOffice)...")
        def thread_odt():
            docx_path = odt_to_docx(filepath)
            if not docx_path:
                root.after(0, lambda: update_progress(0, "Échec conversion ODT"))
                root.after(0, lambda: messagebox.showerror(
                    "Erreur ODT",
                    "LibreOffice non trouvé ou conversion échouée.\n"
                    "Installez LibreOffice :\n"
                    "sudo apt install libreoffice"
                ))
                return
            root.after(0, lambda: continue_processing(docx_path))
        threading.Thread(target=thread_odt, daemon=True).start()
        return

    # --- DOCX direct ---
    elif ext != ".docx":
        update_progress(0, "Format non supporté.")
        messagebox.showwarning("Format non supporté", "Seuls PDF, DOCX et ODT sont acceptés.")
        return

    # --- DOCX direct (sans conversion) ---
    continue_processing(filepath)


# -------------------------------------------------
# Sélection de fichier
# -------------------------------------------------
def select_file():
    filepath = filedialog.askopenfilename(
        title="Sélectionner un document",
        filetypes=[
            ("Documents", "*.pdf *.docx *.odt"),
            ("PDF", "*.pdf"),
            ("Word", "*.docx"),
            ("OpenDocument", "*.odt"),
            ("Tous les fichiers", "*.*")
        ]
    )
    if filepath:
        handle_file(filepath)


# -------------------------------------------------
# Création interface (2 arguments seulement)
# -------------------------------------------------
create_interface(root, select_file)


# -------------------------------------------------
# Vérification Adobe au démarrage
# -------------------------------------------------
check_adobe_status()


# -------------------------------------------------
# Lancement application
# -------------------------------------------------
root.mainloop()
