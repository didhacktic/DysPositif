# converters/pdf_to_docx.py
"""
Conversion PDF -> DOCX avec callback de progression optionnel.

Signature recommandée :
    pdf_to_docx(pdf_path: str, progress_callback: Optional[Callable[[int, str], None]] = None) -> str | None

Le convertisseur appelle progress_callback(percent, message) aux étapes clefs si un callback est fourni.
Il retourne le chemin du .docx produit (ou None en cas d'erreur).
"""
from typing import Optional, Callable
import os
import json
import traceback

CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "../utils/pdfservices-api-credentials.json")

def pdf_to_docx(pdf_path: str, progress_callback: Optional[Callable[[int, str], None]] = None) -> str | None:
    def _prog(p: int, msg: str):
        if progress_callback:
            try:
                progress_callback(int(p), str(msg))
            except Exception:
                # Ne jamais laisser le callback casser la conversion
                traceback.print_exc()

    _prog(0, "Initialisation conversion PDF...")
    if not os.path.exists(pdf_path):
        _prog(0, f"Fichier introuvable : {pdf_path}")
        return None

    if not os.path.exists(CREDENTIALS_PATH):
        _prog(0, f"Credentials manquants : {CREDENTIALS_PATH}")
        return None

    try:
        _prog(10, "Connexion Adobe PDF Services...")
        with open(CREDENTIALS_PATH) as f:
            creds = json.load(f)
        CLIENT_ID = creds.get("CLIENT_ID")
        CLIENT_SECRET = creds.get("CLIENT_SECRETS", [None])[0]
        if not CLIENT_ID or not CLIENT_SECRET:
            _prog(0, "Credentials Adobe invalides")
            return None

        # import retardé pour éviter crash si le SDK manque
        from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
        from adobe.pdfservices.operation.pdf_services import PDFServices
        from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
        from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
        from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

        _prog(40, "Upload du PDF...")
        credentials = ServicePrincipalCredentials(CLIENT_ID, CLIENT_SECRET)
        pdf_services = PDFServices(credentials=credentials)

        with open(pdf_path, "rb") as f:
            asset = pdf_services.upload(f.read(), PDFServicesMediaType.PDF)

        _prog(60, "Conversion en cours...")
        job = ExportPDFJob(asset, ExportPDFParams(ExportPDFTargetFormat.DOCX))
        location = pdf_services.submit(job)
        result = pdf_services.get_job_result(location, ExportPDFResult)

        _prog(80, "Téléchargement du résultat...")
        stream = pdf_services.get_content(result.get_result().get_asset())

        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        _prog(90, "Écriture du fichier DOCX...")
        with open(docx_path, "wb") as f:
            f.write(stream.get_input_stream())

        _prog(100, "Conversion terminée")
        return docx_path

    except Exception as e:
        traceback.print_exc()
        _prog(0, f"Échec conversion PDF : {e}")
        return None