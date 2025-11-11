# ui/interface.py
"""
Interface graphique principale – VERSION FINALE 100% FONCTIONNELLE
Barre de progression, erreurs dans la fenêtre, tout propre
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from config.settings import options

# Variables globales pour progression
progress_bar = None
progress_text = None

def show_error(title: str, message: str):
    print(f"[ERREUR GUI] {title}: {message}")
    messagebox.showerror(title, message)
    update_progress(0, f"ERREUR : {message}")

def show_info(title: str, message: str):
    print(f"[INFO GUI] {title}: {message}")
    messagebox.showinfo(title, message)

def update_progress(value: int, text: str):
    global progress_bar, progress_text
    if progress_bar:
        progress_bar['value'] = value
    if progress_text:
        progress_text.config(state="normal")
        progress_text.delete(1.0, "end")
        progress_text.insert("end", text)
        progress_text.see("end")
        progress_text.config(state="disabled")
    if 'root' in globals() and root:
        root.update_idletasks()

# Variable globale pour la fenêtre principale
root = None

def create_interface(main_root, callback):
    global root, progress_bar, progress_text
    root = main_root
    root.title("Dys’Positif – Adaptation pour la dyslexie")
    root.geometry("860x1120")
    root.configure(bg="#f5f5f5")
    root.resizable(False, False)

    # === Mise à jour des options par défaut ===
    options.update({
        'police': tk.StringVar(value="OpenDyslexic"),
        'taille': tk.IntVar(value=16),
        'interligne': tk.BooleanVar(value=True),
        'espacement': tk.BooleanVar(value=True),
        'syllabes': tk.BooleanVar(value=False),  # DÉSACTIVÉ PAR DÉFAUT
        'griser_muettes': tk.BooleanVar(value=False),  # DÉSACTIVÉ PAR DÉFAUT
        'multicolore': tk.BooleanVar(value=False),
        'position': tk.BooleanVar(value=False),
        'format': tk.StringVar(value="A3"),
        'agrandir_objets': tk.BooleanVar(value=True),
    })
    v = options

    # === En-tête ===
    tk.Label(root, text="Dys’Positif", font=("Helvetica", 32, "bold"), bg="#f5f5f5", fg="#2c3e50").pack(pady=20)
    tk.Label(root, text="PDF • DOCX → Document adapté dyslexie", font=("Helvetica", 13), bg="#f5f5f5", fg="#555555").pack(pady=5)

    # === Bouton central ===
    tk.Button(
        root,
        text="SÉLECTIONNER UN FICHIER\n(PDF ou DOCX)",
        font=("Helvetica", 18, "bold"),
        bg="#0066cc", fg="white",
        width=38, height=3,
        relief="raised", bd=6,
        command=callback
    ).pack(pady=(20, 35))

    # === Boîte des paramètres ===
    params_frame = tk.Frame(root, bg="white", relief="groove", bd=8, padx=35, pady=15)
    params_frame.pack(pady=10, padx=50, fill="both", expand=True)

    # Police
    tk.Label(params_frame, text="Police", font=("Helvetica", 12, "bold"), bg="white").pack(anchor="w", pady=(0, 4))
    ttk.Combobox(
        params_frame,
        textvariable=v['police'],
        values=["OpenDyslexic", "Comic Sans MS", "Arial", "Verdana"],
        state="readonly",
        width=30,
        font=("Helvetica", 11)
    ).pack(anchor="w", pady=(0, 10))

   # Taille
    taille_frame = tk.Frame(params_frame, bg="white")
    taille_frame.pack(anchor="w", pady=(0, 10))
    tk.Label(taille_frame, text="Taille :", font=("Helvetica", 12, "bold"), bg="white").pack(side="left")
    tk.Scale(
        taille_frame,
        from_=12, to=32,
        orient="horizontal",
        variable=v['taille'],
        length=380,
        showvalue=True,
        tickinterval=4,
        font=("Helvetica", 10)
    ).pack(side="left", padx=(10, 0))

    # Espacement
    tk.Label(params_frame, text="Espacement", font=("Helvetica", 12, "bold"), bg="white").pack(anchor="w", pady=(5, 4))
    tk.Checkbutton(params_frame, text="Espacement entre lettres : 2,4 pt", variable=v['espacement'], bg="white", font=("Helvetica", 11)).pack(anchor="w", pady=3)
    tk.Checkbutton(params_frame, text="Interlignage : 1,5", variable=v['interligne'], bg="white", font=("Helvetica", 11)).pack(anchor="w", pady=3)

    # Couleurs des lettres
    tk.Label(params_frame, text="Couleurs des lettres", font=("Helvetica", 12, "bold"), bg="white").pack(anchor="w", pady=(10, 4))
    tk.Checkbutton(params_frame, text="Coloration syllabique (LireCouleur)", variable=v['syllabes'], bg="white", font=("Helvetica", 11)).pack(anchor="w", pady=3)
    tk.Checkbutton(params_frame, text="Griser les syllabes muettes (e, s, ent, ez...)", variable=v['griser_muettes'], bg="white", font=("Helvetica", 11)).pack(anchor="w", pady=3)

    # Coloration des nombres
    tk.Label(params_frame, text="Coloration des nombres", font=("Helvetica", 12, "bold"), bg="white").pack(anchor="w", pady=(10, 4))
    tk.Checkbutton(params_frame, text="Par position (u=bleu, d=rouge, c=vert)", variable=v['position'], bg="white", font=("Helvetica", 11)).pack(anchor="w", padx=25, pady=3)
    tk.Checkbutton(params_frame, text="Multicolore", variable=v['multicolore'], bg="white", font=("Helvetica", 11)).pack(anchor="w", padx=25, pady=3)

    # Synchronisation des modes couleur
    def sync_color_modes(*args):
        if v['multicolore'].get():
            v['position'].set(False)
        elif v['position'].get():
            v['multicolore'].set(False)
    v['multicolore'].trace("w", sync_color_modes)
    v['position'].trace("w", sync_color_modes)

    # Format de sortie
    tk.Label(params_frame, text="Format de sortie", font=("Helvetica", 12, "bold"), bg="white").pack(anchor="w", pady=(10, 4))
    format_frame = tk.Frame(params_frame, bg="white")
    format_frame.pack(anchor="w", pady=0)
    ttk.Combobox(
        format_frame,
        textvariable=v['format'],
        values=["A3", "A4"],
        state="readonly",
        width=12,
        font=("Helvetica", 11)
    ).pack(side="left")
    tk.Checkbutton(
        format_frame,
        text="Agrandir tableaux et zones de texte en A3 (+40 %)",
        variable=v['agrandir_objets'],
        bg="white",
        font=("Helvetica", 11)
    ).pack(side="left", padx=(20, 0))

    # === BARRE DE PROGRESSION ===
    progress_bar = ttk.Progressbar(root, mode="determinate", maximum=100, length=680)
    progress_bar.pack(pady=15)
    progress_text = scrolledtext.ScrolledText(
        root,
        height=6,
        font=("Helvetica", 12),
        bg="white",
        fg="#2c3e50",
        wrap="word",
        state="disabled"
    )
    progress_text.pack(pady=10, padx=80, fill="both", expand=True)

    # === BOUTON QUITTER ===
    footer = tk.Frame(root, bg="#f5f5f5")
    footer.pack(side="bottom", pady=25)
    btn_quit = tk.Button(
        footer,
        text="QUITTER",
        font=("Helvetica", 14, "bold"),
        bg="#ecf0f1", fg="#2c3e50",
        activebackground="#e74c3c", activeforeground="white",
        width=18, height=2,
        relief="raised", bd=4,
        command=root.quit
    )
    btn_quit.pack(pady=12)
    btn_quit.bind("<Enter>", lambda e: btn_quit.config(bg="#e74c3c", fg="white"))
    btn_quit.bind("<Leave>", lambda e: btn_quit.config(bg="#ecf0f1", fg="#2c3e50"))

    # === INITIALISATION ===
    update_progress(0, "Prêt à traiter un document.")