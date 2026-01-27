import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    DECIPHER_BASE = os.getenv("DECIPHER_BASE")
    DECIPHER_API_KEY = os.getenv("DECIPHER_API_KEY")
    FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "dev")

    @classmethod
    def validate(cls):
        missing = []
        if not cls.DECIPHER_BASE:
            missing.append("DECIPHER_BASE")
        if not cls.DECIPHER_API_KEY:
            missing.append("DECIPHER_API_KEY")

        if missing:
            raise RuntimeError(
                "Missing required environment variables: " + ", ".join(missing)
            )
