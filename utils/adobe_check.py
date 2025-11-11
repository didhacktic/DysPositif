# utils/adobe_check.py
import os

CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "pdfservices-api-credentials.json")
ADOBE_MISSING = not os.path.exists(CREDENTIALS_PATH)