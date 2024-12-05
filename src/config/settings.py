import os
from dotenv import load_dotenv

load_dotenv

class Settings:
    URL_RPA_CHALLENGE = os.getenv("URL_RPA_CHALLENGE")
    EXCEL_FILE = os.getenv("EXCEL_FILE")

settings = Settings()