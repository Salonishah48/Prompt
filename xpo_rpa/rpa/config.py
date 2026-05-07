"""
Configuration loader - reads credentials and paths from input_data folder.
"""

import json
import os
from pathlib import Path


class Config:
    # Automatically find correct base directory
    BASE_DIR = Path(__file__).parent.parent
    INPUT_DIR = BASE_DIR / "input_data"
    DOWNLOAD_DIR = BASE_DIR / "downloads"
    LOG_DIR = BASE_DIR / "logs"

    XPO_URL = "https://ext-web.ltl-xpo.com/app/home/shipments"
    BATCH_SIZE = 10

    def __init__(self):
        self.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
        self.LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.INPUT_DIR.mkdir(parents=True, exist_ok=True)

        print(f"[CONFIG] Looking for credentials at: {self.INPUT_DIR / 'credentials.json'}")

        self.credentials = self._load_credentials()

    def _load_credentials(self):
        cred_file = self.INPUT_DIR / "credentials.json"

        if not cred_file.exists():
            template = {"login_id": "your_username_here", "password": "your_password_here"}
            with open(cred_file, "w") as f:
                json.dump(template, f, indent=2)
            raise FileNotFoundError(
                f"\n\n*** CREDENTIALS FILE NOT FOUND ***\n"
                f"A template was created at:\n  {cred_file}\n"
                f"Open that file in Notepad, fill in your login_id and password, then run again.\n"
            )

        try:
            with open(cred_file, "r", encoding="utf-8") as f:
                creds = json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"\n\n*** credentials.json HAS A FORMAT ERROR ***\n"
                f"File: {cred_file}\n"
                f"Error: {e}\n\n"
                f"Make sure the file looks exactly like this:\n"
                f'{{\n  "login_id": "your_username",\n  "password": "your_password"\n}}\n'
                f"Use straight quotes \" not curly quotes.\n"
            )

        print(f"[CONFIG] Credentials loaded. login_id = '{creds.get('login_id', 'MISSING')}'")
        return creds

    def validate(self):
        required = ["login_id", "password"]
        for key in required:
            val = self.credentials.get(key, "")
            if not val:
                raise ValueError(f"\n*** '{key}' is empty in credentials.json. Please fill it in. ***\n")
            if val in ("your_username_here", "your_password_here", "your_actual_username", "your_actual_password"):
                raise ValueError(
                    f"\n*** credentials.json still has placeholder text for '{key}' ***\n"
                    f"Open input_data/credentials.json and replace the placeholder with your real {key}.\n"
                )
        print(f"[CONFIG] Credentials validated OK.")
        return True

    @property
    def login_id(self):
        return self.credentials["login_id"]

    @property
    def password(self):
        return self.credentials["password"]
