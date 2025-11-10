# -------------------------------------------------
# utils/adobe_check.py – Vérification FORCÉE au démarrage (NE PLUS RIEN AVALER)
# -------------------------------------------------

import os
import json
import traceback
from tkinter import messagebox

def check_adobe_status():
    """
    Vérifie que le SDK Adobe fonctionne IMMÉDIATEMENT au démarrage.
    LÈVE UNE EXCEPTION SI QUELQUE CHOSE NE VA PAS.
    """
    credentials_path = os.path.join(os.path.dirname(__file__), "pdfservices-api-credentials.json")

    if not os.path.exists(credentials_path):
        error = f"[Dys’Positif] FICHIER CREDENTIALS MANQUANT :\n{credentials_path}\n\n→ Conversion PDF désactivée."
        messagebox.showwarning("Adobe PDF Services", error)
        raise FileNotFoundError(error)

    try:
        with open(credentials_path) as f:
            creds = json.load(f)
        client_id = creds.get("CLIENT_ID")
        client_secret = creds.get("CLIENT_SECRETS", [None])[0]

        if not client_id or not client_secret:
            raise ValueError("CLIENT_ID ou CLIENT_SECRETS vide ou manquant")

        # TEST RÉEL DU SDK
        from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
        from adobe.pdfservices.operation.pdf_services import PDFServices

        credentials = ServicePrincipalCredentials(client_id, client_secret)
        PDFServices(credentials=credentials)  # Test connexion

        print("[Dys’Positif] Adobe PDF Services : CONNECTÉ ✓")
        return True

    except Exception as e:
        error = f"[Dys’Positif] ÉCHEC CONNEXION ADOBE :\n{e}\n\n{traceback.format_exc()}"
        messagebox.showerror("ERREUR ADOBE", error)
        print(error)
        raise  # ← CRUCIAL : on laisse l’exception remonter → excepthook l’attrape