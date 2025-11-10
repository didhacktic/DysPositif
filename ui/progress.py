# converters/pdf_to_docx.py – Avec progression + erreurs dans GUI
import os
import json
from ui.interface import update_progress, show_error

CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "../utils/pdfservices-api-credentials.json")

def pdf_to_docx(pdf_path: str) -> str | None:
    update_progress(10, "Connexion à Adobe PDF Services...")
    
    if not os.path.exists(CREDENTIALS_PATH):
        show_error("Credentials manquants", f"Fichier non trouvé :\n{CREDENTIALS_PATH}")
        return None

    try:
        with open(CREDENTIALS_PATH) as f:
            creds = json.load(f)
        CLIENT_ID = creds["CLIENT_ID"]
        CLIENT_SECRET = creds["CLIENT_SECRETS"][0]
    except Exception as e:
        show_error("Erreur credentials", f"Impossible de lire le fichier JSON :\n{e}")
        return None

    try:
        from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
        from adobe.pdfservices.operation.pdf_services import PDFServices
        from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
        from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
        from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

        update_progress(20, "Authentification Adobe...")
        credentials = ServicePrincipalCredentials(CLIENT_ID, CLIENT_SECRET)
        pdf_services = PDFServices(credentials=credentials)

        update_progress(40, "Upload du PDF...")
        with open(pdf_path, "rb") as f:
            asset = pdf_services.upload(f.read(), PDFServicesMediaType.PDF)

        update_progress(60, "Conversion en cours...")
        job = ExportPDFJob(asset, ExportPDFParams(ExportPDFTargetFormat.DOCX))
        location = pdf_services.submit(job)
        result = pdf_services.get_job_result(location, ExportPDFResult)
        stream = pdf_services.get_content(result.get_result().get_asset())

        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        update_progress(80, "Écriture du fichier DOCX...")
        with open(docx_path, "wb") as f:
            f.write(stream.get_input_stream())

        update_progress(90, "Conversion terminée !")
        return docx_path

    except Exception as e:
        show_error("Échec conversion PDF", f"Adobe a renvoyé une erreur :\n{e}")
        return None