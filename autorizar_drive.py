import os
from google_auth_oauthlib.flow import InstalledAppFlow

BASE = os.path.dirname(os.path.abspath(__file__))

# Incluye el conector de Google Drive y el conector de envío de correos Gmail API
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/gmail.send"
]

secrets = next((os.path.join(BASE, f) for f in ("client_secrets.json", "client_secrets.json.json")
                if os.path.exists(os.path.join(BASE, f))), None)
if not secrets:
    raise SystemExit("No encuentro client_secrets.json en esta carpeta.")

creds = InstalledAppFlow.from_client_secrets_file(secrets, SCOPES).run_local_server(port=0)
with open(os.path.join(BASE, "token.json"), "w") as f:
    f.write(creds.to_json())
print("Listo: token.json actualizado con permisos de Google Drive y Gmail API.")