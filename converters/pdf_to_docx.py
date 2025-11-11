# converters/pdf_to_docx.py
import os
import json
import traceback

CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "../utils/pdfservices-api-credentials.json")

def pdf_to_docx(pdf_path: str) -> str:
    if not os.path.exists(CREDENTIALS_PATH):
        raise FileNotFoundError(f"credentials.json absent → {CREDENTIALS_PATH}")

    with open(CREDENTIALS_PATH) as f:
        creds = json.load(f)
    CLIENT_ID = creds["CLIENT_ID"]
    CLIENT_SECRET = creds["CLIENT_SECRETS"][0]

    from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
    from adobe.pdfservices.operation.pdf_services import PDFServices
    from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
    from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
    from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
    from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
    from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

    credentials = ServicePrincipalCredentials(CLIENT_ID, CLIENT_SECRET)
    pdf_services = PDFServices(credentials=credentials)

    with open(pdf_path, "rb") as f:
        asset = pdf_services.upload(f.read(), PDFServicesMediaType.PDF)

    job = ExportPDFJob(asset, ExportPDFParams(ExportPDFTargetFormat.DOCX))
    location = pdf_services.submit(job)
    result = pdf_services.get_job_result(location, ExportPDFResult)
    stream = pdf_services.get_content(result.get_result().get_asset())

    docx_path = os.path.splitext(pdf_path)[0] + ".docx"
    with open(docx_path, "wb") as f:
        f.write(stream.get_input_stream())

    return docx_path