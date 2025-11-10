# -------------------------
# converters/pdf_to_docx.py 
# -------------------------
import os
import json
import sys
import traceback

print("\n[DEBUG] converters/pdf_to_docx.py CHARGÉ")

CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "../utils/pdfservices-api-credentials.json")
print(f"[DEBUG] Chemin credentials : {CREDENTIALS_PATH}")

def pdf_to_docx(pdf_path: str) -> str:
    print(f"\n[DEBUG] → pdf_to_docx appelé avec : {pdf_path}")
    
    if not os.path.exists(CREDENTIALS_PATH):
        error = f"ERREUR : credentials.json ABSENT → {CREDENTIALS_PATH}"
        print(error)
        traceback.print_stack()
        raise FileNotFoundError(error)

    print("[DEBUG] credentials.json trouvé")

    try:
        with open(CREDENTIALS_PATH) as f:
            creds = json.load(f)
        CLIENT_ID = creds["CLIENT_ID"]
        CLIENT_SECRET = creds["CLIENT_SECRETS"][0]
        print(f"[DEBUG] Credentials chargés : CLIENT_ID = {CLIENT_ID[:10]}...")
    except Exception as e:
        print("ERREUR LECTURE CREDENTIALS")
        traceback.print_exc()
        raise

    print("[DEBUG] Import SDK Adobe...")
    try:
        from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
        from adobe.pdfservices.operation.pdf_services import PDFServices
        from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
        from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
        from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
        from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
        print("[DEBUG] SDK Adobe importé")
    except Exception as e:
        print("ERREUR IMPORT SDK ADOBE")
        traceback.print_exc()
        raise

    try:
        print("[DEBUG] Connexion à Adobe...")
        credentials = ServicePrincipalCredentials(CLIENT_ID, CLIENT_SECRET)
        pdf_services = PDFServices(credentials=credentials)

        print("[DEBUG] Upload du PDF...")
        with open(pdf_path, "rb") as f:
            asset = pdf_services.upload(f.read(), PDFServicesMediaType.PDF)

        print("[DEBUG] Soumission job...")
        job = ExportPDFJob(asset, ExportPDFParams(ExportPDFTargetFormat.DOCX))
        location = pdf_services.submit(job)

        print("[DEBUG] Attente résultat...")
        result = pdf_services.get_job_result(location, ExportPDFResult)
        stream = pdf_services.get_content(result.get_result().get_asset())

        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        print(f"[DEBUG] Écriture → {docx_path}")
        with open(docx_path, "wb") as f:
            f.write(stream.get_input_stream())

        print("[DEBUG] CONVERSION RÉUSSIE")
        return docx_path

    except Exception as e:
        print("\n" + "="*60)
        print("ÉCHEC CONVERSION ADOBE — EXCEPTION CAPTURÉE")
        print("="*60)
        traceback.print_exc()
        print("="*60)
        raise  # ← ON FORCE LE RAISE
