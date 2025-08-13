# imap_reader.py
import imaplib
import email
from email.header import decode_header
import re, hashlib

IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

def _dec(h):
    if not h:
        return ""
    parts = decode_header(h)
    out = []
    for t, enc in parts:
        if isinstance(t, bytes):
            out.append(t.decode(enc or "utf-8", errors="ignore"))
        else:
            out.append(t)
    return "".join(out)

def _get_text_from_email(msg):
    # Prioriza text/plain; si no hay, intenta text/html (sin limpiar)
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if ctype == "text/plain" and "attachment" not in disp:
                return part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="ignore"
                )
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="ignore"
                )
    payload = msg.get_payload(decode=True)
    if payload:
        return payload.decode(msg.get_content_charset() or "utf-8", errors="ignore")
    return ""

def conectar_a_imap(email_addr: str, app_password: str):
    print("-> Iniciando conexión IMAP...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    print("-> Conexión SSL establecida.")
    mail.login(email_addr, app_password)
    print("-> Login exitoso.")
    mail.select("inbox")  # si querés solo lectura: mail.select("inbox", readonly=True)
    print("✅ ¡Conectado exitosamente a IMAP!")
    return mail

def obtener_mails_bancarios(imap, remitentes, limite_por_remitente=10):
    """
    Devuelve lista de dicts: {seq, uid, message_id, id, from, subject, date, body}
    - id = message_id si existe; si no, uid; si no, hash(date+from+subject+body[:40])
    """
    resultados = []

    for remit in remitentes:
        print(f"[*] Buscando correos de: {remit}")
        status, data = imap.search(None, f'(FROM "{remit}")')
        if status != "OK" or not data or not data[0]:
            print("   (no se encontraron mensajes)")
            continue

        seq_ids = data[0].split()[-limite_por_remitente:]  # últimos N

        for seq in seq_ids:
            # UID
            uid = None
            st_uid, data_uid = imap.fetch(seq, "(UID)")
            if st_uid == "OK" and data_uid and data_uid[0]:
                m = re.search(rb"UID (\d+)", data_uid[0])
                if m:
                    uid = m.group(1).decode()

            # Mensaje completo (peek para no marcar leído)
            st_fetch, fetched = imap.fetch(seq, "(BODY.PEEK[])")
            if st_fetch != "OK" or not fetched or not fetched[0]:
                continue

            raw = fetched[0][1]
            msg = email.message_from_bytes(raw)

            message_id = _dec(msg.get("Message-Id")) or _dec(msg.get("Message-ID"))  # variantes
            from_      = _dec(msg.get("From"))
            subject    = _dec(msg.get("Subject"))
            date_      = _dec(msg.get("Date"))
            body       = _get_text_from_email(msg)

            # id robusto
            if message_id:
                stable_id = message_id.strip()
            elif uid:
                stable_id = f"UID:{uid}"
            else:
                sig = f"{date_}|{from_}|{subject}|{(body or '')[:40]}"
                stable_id = "HASH:" + hashlib.sha1(sig.encode("utf-8", errors="ignore")).hexdigest()

            resultados.append({
                "seq": seq.decode(),
                "uid": uid,
                "message_id": message_id,
                "id": stable_id,      # <- usar este en Sheets
                "from": from_,
                "subject": subject,
                "date": date_,
                "body": body,
            })

        print(f"   (+) Encontrados {len(seq_ids)} correos de {remit}")

    print(f"[OK] Total recolectados: {len(resultados)}")
    return resultados
