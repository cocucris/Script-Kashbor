# sheets_writer.py
import os
from googleapiclient.discovery import build
from google.oauth2 import service_account

SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1V4mm0T9tS0tv9BPMFbFObQoRzPGxQ6Hg30jgYoJ6-RM")
SHEET_RANGE = os.getenv("SHEET_RANGE", "AGOSTO!A:G ")  # ej: "AGOSTO!A:G "

HEADERS_ES = [
    "fecha_procesado",       # fecha en que el script procesa el mail
    "remitente",             # dirección de email que envió el mensaje
    "asunto",                # subject del correo
    "monto",                 # cantidad en números
    "tipo_movimiento",       # débito / crédito
    "moneda",                # PYG, USD, etc.
    "id_mensaje",            # identificador único del mail
]

def _get_sheet_name(sheet_range: str) -> str:
    # "Hoja 1!A1" -> "Hoja 1"
    return sheet_range.split("!", 1)[0] if "!" in sheet_range else "Hoja 1"

def conectar_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_JSON, scopes=scopes
    )
    return build('sheets', 'v4', credentials=creds)

def _ensure_headers(svc, spreadsheet_id):
    """Escribe encabezados si A1 está vacío."""
    sheet_name = _get_sheet_name(SHEET_RANGE)
    resp = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1:G1"
    ).execute()

    values = resp.get("values", [])
    if not values or not any(values[0]):
        svc.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": [HEADERS_ES]}
        ).execute()
        print("[OK] Encabezados creados.")

def append_rows(sheets_service, rows):
    """Inserta filas (lista de listas). Asegura encabezados primero."""
    if not rows:
        print("[i] No hay filas para insertar en Sheets.")
        return

    _ensure_headers(sheets_service, SPREADSHEET_ID)

    sheet_name = _get_sheet_name(SHEET_RANGE)
    body = {"values": rows}
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()
    print(f"[OK] Insertadas {len(rows)} filas en Google Sheets.")
