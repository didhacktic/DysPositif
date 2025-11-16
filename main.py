#!/usr/bin/env python3
# main.py
# DEBUG: démarrage main.py
"""
Point d'entrée de l'application avec queue séquentielle et progression via callbacks.

Comportement clé :
- Les convertisseurs (pdf_to_docx / odt_to_docx) acceptent un progress_callback(percent:int, message:str).
- main.py fournit un adapter ui_progress qui poste les updates vers l'UI (root.after -> update_progress).
- Si un batch (plusieurs fichiers) : on empêche l'ouverture individuelle des fichiers et on ouvre
  le dossier DYS final une fois la file terminée.
- Fallback visuel possible : start_progress_busy / stop_progress_busy (si définies dans ui.interface).
"""
import os
import threading
import subprocess
import platform
from tkinter import Tk, messagebox

from ui.interface import create_interface, update_progress
# try to import the optional busy helpers; provide no-op fallback if absent
try:
    from ui.interface import start_progress_busy, stop_progress_busy
except Exception:
    def start_progress_busy(text: str):
        try:
            update_progress(0, text)
        except Exception:
            pass

    def stop_progress_busy(final_text: str | None = None, final_value: int | None = None):
        try:
            if final_text is not None:
                update_progress(final_value or 0, final_text)
        except Exception:
            pass

from converters.pdf_to_docx import pdf_to_docx
from converters.odt_to_docx import odt_to_docx
from core.processor import process_document
from utils.adobe_check import ADOBE_MISSING

# --- Initialisation de la fenêtre principale ---
root = Tk()
root.title("Dys’Positif – Adaptation pour la dyslexie")
root.geometry("900x700")
root.configure(bg="#f8f9fa")
root.resizable(False, False)

if ADOBE_MISSING:
    messagebox.showwarning(
        "Adobe PDF Services désactivé",
        "Fichier manquant :\n"
        "utils/pdfservices-api-credentials.json\n\n"
        "→ La conversion PDF → DOCX ne fonctionnera pas.\n\n"
        "Contactez votre administrateur pour obtenir ce fichier."
    )

# --- Queue séquentielle ---
_files_queue: list[str] = []
_is_processing = False
_queue_lock = threading.Lock()

# Batch handling
_suppress_open_for_batch = False
_last_output_folder = None
_last_output_lock = threading.Lock()

current_file = None

# Adapter progress callback to UI (always post to UI thread)
def ui_progress(percent: int, message: str):
    # percent can be 0..100; ensure int
    try:
        pct = int(percent)
    except Exception:
        pct = 0
    # update UI on main thread
    root.after(0, lambda: update_progress(pct, message))
    # if percent indicates completion, stop busy animation if any
    if pct >= 100:
        try:
            root.after(0, lambda: stop_progress_busy(message, final_value=100))
        except Exception:
            pass

def _open_folder(path: str):
    """Open file manager on path (cross-platform)."""
    try:
        if platform.system() == "Linux":
            subprocess.call(["xdg-open", path])
        elif platform.system() == "Darwin":
            subprocess.call(["open", path])
        else:
            os.startfile(path)
    except Exception as e:
        try:
            root.after(0, lambda: messagebox.showwarning("Ouverture dossier", f"Impossible d'ouvrir le dossier :\n{e}"))
        except Exception:
            pass

def _process_next_in_queue():
    """
    Démarre le traitement du prochain fichier dans la file (si présent).
    """
    global _is_processing, _suppress_open_for_batch, _last_output_folder
    with _queue_lock:
        if not _files_queue:
            _is_processing = False
            if _suppress_open_for_batch:
                # read and reset under lock
                with _last_output_lock:
                    folder = _last_output_folder
                    _last_output_folder = None
                _suppress_open_for_batch = False
                if folder:
                    root.after(0, lambda: update_progress(0, f"Fin batch — ouverture dossier: {folder}"))
                    root.after(0, lambda: _open_folder(folder))
                else:
                    root.after(0, lambda: update_progress(0, "Fin batch — aucun dossier trouvé à ouvrir"))
            else:
                root.after(0, lambda: update_progress(0, "Prêt à traiter un document."))
            return
        filepath = _files_queue.pop(0)
        _is_processing = True

    handle_file(filepath, on_complete=_process_next_in_queue)

def enqueue_files(filepaths):
    """
    Ajoute des fichiers à la queue et déclenche le démarrage si rien n'est en cours.
    filepaths : iterable of paths
    """
    global _suppress_open_for_batch
    incoming_count = 0
    with _queue_lock:
        existing = set(_files_queue)
        for p in filepaths:
            if p not in existing:
                _files_queue.append(p)
                existing.add(p)
                incoming_count += 1
        if incoming_count > 1 or len(_files_queue) > 1:
            _suppress_open_for_batch = True

    root.after(0, lambda: update_progress(0, f"{len(_files_queue)} fichier(s) en file d'attente. batch={_suppress_open_for_batch}"))
    if not _is_processing:
        root.after(0, _process_next_in_queue)

def continue_processing(docx_path: str, on_complete=None, open_after: bool = True):
    """
    Lance process_document(docx_path, open_after=...) dans un thread et mémorise le dossier de sortie.
    """
    # capture open_after
    local_open_after = bool(open_after)

    root.after(0, lambda: update_progress(0, f"continue_processing: {os.path.basename(docx_path)}  open_after={local_open_after}"))

    def _worker():
        global _last_output_folder
        output_path = None
        try:
            # process_document is expected to return output_path (or raise on error)
            output_path = process_document(docx_path, open_after=local_open_after)
        except Exception as e:
            root.after(0, lambda: update_progress(0, "Erreur durant le traitement"))
            root.after(0, lambda: messagebox.showerror("Erreur traitement", str(e)))
        finally:
            if output_path:
                folder = os.path.dirname(output_path)
                with _last_output_lock:
                    _last_output_folder = folder
                root.after(0, lambda f=folder: update_progress(0, f"_last_output_folder défini → {f}"))
                root.after(0, lambda p=os.path.basename(output_path): update_progress(0, f"Fichier traité → {p}"))
            if on_complete:
                root.after(0, on_complete)

    threading.Thread(target=_worker, daemon=True).start()

def handle_file(filepath: str, on_complete=None):
    """
    Gère la conversion initiale (PDF/ODT -> DOCX) puis lance le traitement du .docx.
    """
    global current_file
    current_file = filepath
    ext = os.path.splitext(filepath)[1].lower()
    root.after(0, lambda: update_progress(0, f"handle_file: {os.path.basename(filepath)} (ext={ext})"))

    def _call_on_complete_safe():
        if on_complete:
            try:
                on_complete()
            except Exception:
                pass

    if ext == ".pdf":
        root.after(0, lambda: start_progress_busy("Conversion PDF → DOCX (Adobe)..."))

        def thread_pdf():
            try:
                docx_path = pdf_to_docx(filepath, progress_callback=ui_progress)
                if docx_path:
                    # conversion done; stop busy if not already stopped by progress_callback
                    root.after(0, lambda: stop_progress_busy("Conversion terminée", final_value=30))
                    # pass open_after flag depending on batch mode
                    root.after(0, lambda: continue_processing(docx_path, on_complete=on_complete, open_after=(not _suppress_open_for_batch)))
                else:
                    root.after(0, lambda: stop_progress_busy("Échec conversion PDF", final_value=0))
                    root.after(0, lambda: update_progress(0, "Échec conversion PDF"))
                    root.after(0, lambda: messagebox.showerror("Erreur PDF", "Conversion échouée"))
                    root.after(0, _call_on_complete_safe)
            except Exception as e:
                root.after(0, lambda: stop_progress_busy("Échec conversion PDF", final_value=0))
                root.after(0, lambda: update_progress(0, "Échec conversion PDF"))
                root.after(0, lambda: messagebox.showerror("Erreur PDF", str(e)))
                root.after(0, _call_on_complete_safe)

        threading.Thread(target=thread_pdf, daemon=True).start()
        return

    elif ext == ".odt":
        root.after(0, lambda: start_progress_busy("Conversion ODT → DOCX (LibreOffice)..."))

        def thread_odt():
            try:
                docx_path = odt_to_docx(filepath, progress_callback=ui_progress)
                if docx_path:
                    root.after(0, lambda: stop_progress_busy("Conversion terminée", final_value=30))
                    root.after(0, lambda: continue_processing(docx_path, on_complete=on_complete, open_after=(not _suppress_open_for_batch)))
                else:
                    root.after(0, lambda: stop_progress_busy("Échec conversion ODT", final_value=0))
                    root.after(0, lambda: update_progress(0, "Échec conversion ODT"))
                    root.after(0, lambda: messagebox.showerror("Erreur ODT", "Conversion échouée"))
                    root.after(0, _call_on_complete_safe)
            except Exception as e:
                root.after(0, lambda: stop_progress_busy("Échec conversion ODT", final_value=0))
                root.after(0, lambda: update_progress(0, "Échec conversion ODT"))
                root.after(0, lambda: messagebox.showerror("Erreur ODT", str(e)))
                root.after(0, _call_on_complete_safe)

        threading.Thread(target=thread_odt, daemon=True).start()
        return

    elif ext != ".docx":
        root.after(0, lambda: update_progress(0, "Format non supporté"))
        root.after(0, lambda: messagebox.showwarning("Format non supporté", "Seuls PDF, DOCX et ODT sont acceptés."))
        root.after(0, _call_on_complete_safe)
        return

    # .docx : pas de conversion, lancer traitement
    continue_processing(filepath, on_complete=on_complete, open_after=(not _suppress_open_for_batch))

# Intégration UI
create_interface(root, enqueue_files)

# Lancement de la boucle Tk
root.mainloop()