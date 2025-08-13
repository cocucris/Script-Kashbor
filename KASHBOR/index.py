# index.py
import os
import re
from datetime import datetime
from dotenv import load_dotenv
from imap_reader import conectar_a_imap, obtener_mails_bancarios
from sheets_writer import conectar_sheets, append_rows


# Cargar .env ANTES de importar módulos que leen variables
load_dotenv()



# ====== Parsers simples ======

CTX_WORDS = ("monto", "importe", "acreditad", "transferenc", "pago", "depósito", "deposito", "crédito", "credito", "débito", "debito")

def _to_int_from_mixed(numstr: str) -> int | None:
    """
    '10,000.00' -> 10000
    '1.234.567,89' -> 1234567
    '100000' -> 100000
    También corrige casos como '50000' (sin punto) que son montos grandes.
    """
    s = numstr.strip()
    if not s:
        return None

    # si hay ambos separadores, la última aparición suele ser el decimal
    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        dec_sep = "," if last_comma > last_dot else "."
        s_int = s[:s.rfind(dec_sep)]
        s_int = s_int.replace(".", "").replace(",", "")
    else:
        # sin decimales: quitamos separadores de miles
        s_int = s.replace(".", "").replace(",", "")

    # 📌 Si es un número largo sin separadores (>= 5 dígitos), probablemente es miles en PYG
    if s_int.isdigit() and len(s_int) >= 5:
        # Aquí podrías formatear, pero como lo devolvemos int, no hace falta insertar punto
        return int(s_int)

    return int(s_int) if s_int.isdigit() else None

def extraer_monto(texto: str) -> int | None:
    """
    Devuelve el monto como entero (PYG) soportando:
    - Prefijo de moneda: 'Gs. *****28.000'
    - Sufijo de moneda: '10,000.00 GS.'
    - Miles con . o , y decimales opcionales
    - Evita números de 'cuenta' cercanos
    """
    if not texto:
        return None

    # 1) Prefijo de moneda -> número
    pat_prefix = re.compile(
        r'(?:Gs\.?|GS\.?|₲|PYG|USD|US\$|\$)\s*[^0-9]{0,8}\s*'
        r'([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,]\d{2})?|\d+)',
        re.IGNORECASE
    )
    # 2) Número -> sufijo de moneda
    pat_suffix = re.compile(
        r'([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,]\d{2})?|\d+)\s*'
        r'(?:Gs\.?|GS\.?|₲|PYG|USD|US\$|\$)',
        re.IGNORECASE
    )
    candidates = []

    for m in pat_prefix.finditer(texto):
        val = _to_int_from_mixed(m.group(1))
        if val is not None:
            candidates.append((val, m.start()))

    for m in pat_suffix.finditer(texto):
        val = _to_int_from_mixed(m.group(1))
        if val is not None:
            candidates.append((val, m.start()))

    # 3) Fallback: números grandes con miles o enteros largos (evitando 'cuenta')
    if not candidates:
        for m in re.finditer(r'\b([0-9]{1,3}(?:[.,][0-9]{3})+|\d{4,})\b', texto):
            # evitar números de cuenta si hay 'cuenta' cerca
            prev = texto[max(0, m.start()-20):m.start()].lower()
            if "cuenta" in prev:
                continue
            val = _to_int_from_mixed(m.group(1))
            if val is not None:
                candidates.append((val, m.start()))

    if not candidates:
        return None

    # 4) Rankear por cercanía a palabras de contexto (monto/importe/etc.)
    t_low = texto.lower()
    anchors = []
    for w in CTX_WORDS:
        pos = 0
        while True:
            i = t_low.find(w, pos)
            if i == -1: break
            anchors.append(i); pos = i + len(w)

    if anchors:
        best = min(
            candidates,
            key=lambda c: min(abs(c[1]-a) for a in anchors)
        )
        return best[0]

    # si no hay contexto, devolver el primero “razonable” (ya orden natural)
    return candidates[0][0]
    # Elegimos el más "monto-like": primero el que tenga puntos (miles); si no, el más grande
    con_puntos = [c for c in candidatos if "." in c]
    elegido = con_puntos[0] if con_puntos else max(candidatos, key=lambda s: int(s.replace(".", "")))
    num_str = elegido.replace(".", "").replace(",", "")
    return int(num_str) if num_str.isdigit() else None


def detectar_moneda(texto: str) -> str:
    t = (texto or "").upper()
    if any(x in t for x in ("USD", "US$", "$")):
        return "USD"
    return "PYG"

def inferir_tipo_movimiento(texto: str) -> str:
    t = (texto or "").lower()
    if any(w in t for w in ["débito","debito","compra","consumo","pago","retiro","extracción","extraccion","gasto", "pago enviado", "gasto enviado", "pago realizado", "gasto realizado", "pago efectuado", "gasto efectuado","enviado","aviso de transferencia enviada"]):
        return "debito"
    if any(w in t for w in ["abono","acreditación","acreditacion","depósito","deposito","transferencia recibida","pago recibido","crédito","credito", "pago acreditado", "gasto acreditado", "pago recibido", "gasto recibido","recibido","devolucion", "devolución","reembolso","reintegro","reintegro recibido"]):
        return "credito"
    return "desconocido"

def cargar_ids_existentes(sheets_service) -> set:
    """Lee la columna id_mensaje (G) y devuelve un set con los IDs ya insertados."""
    import os
    sheet_range = os.getenv("SHEET_RANGE", "Hoja 1!A1")
    sheet_name = sheet_range.split("!", 1)[0] if "!" in sheet_range else "Hoja 1"
    resp = sheets_service.spreadsheets().values().get(
        spreadsheetId=os.getenv("SPREADSHEET_ID"),
        range=f"{sheet_name}!G2:G"
    ).execute()
    values = resp.get("values", [])
    return {row[0] for row in values if row}  # set de ids


# ====== Main ======
def main():
    email_user = os.getenv("EMAIL_USER")
    email_pass = os.getenv("EMAIL_PASS")
    remitentes = [r.strip() for r in os.getenv("REMITENTES_BANCOS", "").split(",") if r.strip()]

    if not email_user or not email_pass:
        print("❌ Falta EMAIL_USER o EMAIL_PASS en .env"); return
    if not remitentes:
        print("❌ REMITENTES_BANCOS vacío en .env"); return

    imap = conectar_a_imap(email_user, email_pass)
    mails = obtener_mails_bancarios(imap, remitentes, limite_por_remitente=10)
    print(f"[i] Mails obtenidos: {len(mails)}")

    # 3) Preparar filas con dedupe por id_mensaje
    ahora = "'" + datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # texto legible en Sheets
    sheets = conectar_sheets()
    ids_existentes = cargar_ids_existentes(sheets)

    rows = []
    saltados = 0
    for m in mails:
        id_msg = m.get("id", "")
        if id_msg in ids_existentes:
            saltados += 1
            continue

        texto = f"{m.get('subject','')} {m.get('body','')}"
        monto    = extraer_monto(texto) or 0
        moneda   = detectar_moneda(texto)
        tipo_mov = inferir_tipo_movimiento(texto)

        rows.append([
            ahora,                 # fecha_procesado (texto)
            m.get("from", ""),     # remitente
            m.get("subject", ""),  # asunto
            monto,                 # monto (int)
            tipo_mov,              # tipo_movimiento
            moneda,                # moneda
            id_msg,                # id_mensaje (estable)
        ])

    if rows:
        append_rows(sheets, rows)
        print(f"[OK] Insertadas {len(rows)} nuevas filas. Omitidas por duplicado: {saltados}.")
    else:
        print(f"[i] No hay filas para insertar. Omitidas por duplicado: {saltados}.")

'''if __name__ == "__main__":
    main()'''

if __name__ == "__main__":
    import time
    from datetime import datetime

    while True:
        try:
            main()
        except Exception as e:
            print(f"Error: {e}")
        time.sleep(10)
