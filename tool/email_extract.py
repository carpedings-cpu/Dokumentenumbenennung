"""
E-Mail-Extraktion fuer den Dokumenten-Umbenenner.

Beim Reinziehen einer E-Mail (.eml oder .msg) wird daraus erzeugt:
  - der Mailtext (inkl. Kopfzeilen) als PDF,
  - jede angehaengte Datei als eigene Datei.

Die erzeugten Dateien werden anschliessend wie normale Dokumente nach der
Verfahrensanweisung benannt.

.eml wird mit der Standardbibliothek gelesen. .msg (Outlook) nutzt das Paket
'extract-msg' (in der .exe enthalten; aus dem Quellcode ggf. `pip install
extract-msg`).
"""

import html
import os
import re
from email import policy
from email.parser import BytesParser


def ist_email(pfad):
    return os.path.splitext(pfad)[1].lower() in (".eml", ".msg")


# ----------------------------------------------------------------- Hilfen
def _html_zu_text(h):
    if not h:
        return ""
    h = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", h)
    h = re.sub(r"(?i)<br\s*/?>", "\n", h)
    h = re.sub(r"(?i)</p\s*>", "\n\n", h)
    h = re.sub(r"<[^>]+>", " ", h)
    h = html.unescape(h)
    h = re.sub(r"[ \t]+", " ", h)
    h = re.sub(r"\n\s*\n\s*\n+", "\n\n", h)
    return h.strip()


def _sicherer_dateiname(name):
    name = os.path.basename(str(name or "")).strip() or "Anhang"
    return re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name)


def _freier_name(ordner, stem, ext):
    kandidat = f"{stem}{ext}"
    i = 2
    while os.path.exists(os.path.join(ordner, kandidat)):
        kandidat = f"{stem}-{i:02d}{ext}"
        i += 1
    return kandidat


# ----------------------------------------------------------------- Lesen
def _eml_lesen(pfad):
    with open(pfad, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    kopf = {"Von": msg.get("From", ""), "An": msg.get("To", ""),
            "Datum": msg.get("Date", ""), "Betreff": msg.get("Subject", "")}

    plain = htmlteil = None
    anhaenge = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        disp = part.get_content_disposition()
        ctype = part.get_content_type()
        cid = part.get("Content-ID")
        if disp == "attachment" or part.get_filename():
            data = part.get_payload(decode=True) or b""
            inline = (disp == "inline") or bool(cid)
            anhaenge.append((part.get_filename() or "Anhang", data, ctype, inline))
        elif ctype == "text/plain" and plain is None:
            plain = part.get_content()
        elif ctype == "text/html" and htmlteil is None:
            htmlteil = part.get_content()
    body = plain if plain else _html_zu_text(htmlteil)
    return kopf, body or "", anhaenge


def _msg_lesen(pfad):
    try:
        import extract_msg
    except ImportError as e:
        raise RuntimeError("Für .msg-Dateien wird 'extract-msg' benötigt "
                            "(in der .exe enthalten; sonst: pip install extract-msg).") from e
    m = extract_msg.Message(pfad)
    try:
        kopf = {"Von": m.sender or "", "An": m.to or "",
                "Datum": str(m.date or ""), "Betreff": m.subject or ""}
        body = m.body or ""
        if not body and getattr(m, "htmlBody", None):
            hb = m.htmlBody
            if isinstance(hb, bytes):
                hb = hb.decode("utf-8", "replace")
            body = _html_zu_text(hb)
        anhaenge = []
        for att in m.attachments:
            name = att.longFilename or att.shortFilename or "Anhang"
            data = att.data
            if isinstance(data, str):
                data = data.encode("utf-8", "replace")
            if not data:
                continue
            cid = getattr(att, "cid", None) or getattr(att, "contentId", None)
            ctype = getattr(att, "mimetype", "") or ""
            inline = bool(cid)
            anhaenge.append((name, data, ctype, inline))
        return kopf, body or "", anhaenge
    finally:
        try:
            m.close()
        except Exception:
            pass


# ----------------------------------------------------------------- PDF
def _pdf_escape(s):
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _umbrechen(text, breite=90):
    zeilen = []
    for absatz in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not absatz:
            zeilen.append("")
            continue
        cur = ""
        for wort in absatz.split(" "):
            while len(wort) > breite:           # ueberlange Woerter hart trennen
                if cur:
                    zeilen.append(cur); cur = ""
                zeilen.append(wort[:breite]); wort = wort[breite:]
            if len(cur) + len(wort) + 1 <= breite:
                cur = (cur + " " + wort).strip()
            else:
                if cur:
                    zeilen.append(cur)
                cur = wort
        zeilen.append(cur)
    return zeilen


def _text_zu_pdf(kopf, body, zielpfad):
    leading, size, left, top, pro_seite = 14, 10, 56, 788, 52
    zeilen = []
    for k in ("Von", "An", "Datum", "Betreff"):
        zeilen += _umbrechen(f"{k}: {kopf.get(k, '')}", 90)
    zeilen += ["", "-" * 80, ""]
    zeilen += _umbrechen(body, 90)

    seiten = [zeilen[i:i + pro_seite] for i in range(0, max(len(zeilen), 1), pro_seite)] or [[""]]

    page_ids, content_ids, nid = [], [], 4
    for _ in seiten:
        page_ids.append(nid); content_ids.append(nid + 1); nid += 2

    objs = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        2: ("<< /Type /Pages /Count %d /Kids [%s] >>"
            % (len(seiten), " ".join(f"{p} 0 R" for p in page_ids))).encode(),
        3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    }
    for idx, seite in enumerate(seiten):
        pid, cid = page_ids[idx], content_ids[idx]
        teile = [b"BT", f"/F1 {size} Tf".encode(), f"{leading} TL".encode(),
                 f"{left} {top} Td".encode()]
        for j, ln in enumerate(seite):
            if j:
                teile.append(b"T*")
            teile.append(b"(" + _pdf_escape(ln).encode("cp1252", "replace") + b") Tj")
        teile.append(b"ET")
        stream = b"\n".join(teile)
        objs[cid] = b"<< /Length %d >>\nstream\n%s\nendstream" % (len(stream), stream)
        objs[pid] = (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
                     b"/Resources << /Font << /F1 3 0 R >> >> /Contents %d 0 R >>" % cid)

    buf = bytearray(b"%PDF-1.4\n")
    offsets = {}
    for i in range(1, nid):
        offsets[i] = len(buf)
        buf += f"{i} 0 obj\n".encode() + objs[i] + b"\nendobj\n"
    xref = len(buf)
    buf += f"xref\n0 {nid}\n".encode() + b"0000000000 65535 f \n"
    for i in range(1, nid):
        buf += f"{offsets[i]:010d} 00000 n \n".encode()
    buf += f"trailer\n<< /Size {nid} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode()
    with open(zielpfad, "wb") as f:
        f.write(buf)


# ----------------------------------------------------------------- API
_BILD_ENDUNGEN = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg",
                  ".emf", ".wmf", ".ico", ".tif", ".tiff")


def _ist_echter_anhang(name, data, ctype, inline):
    """True = echter Datei-Anhang (PDF, Word, Excel, Foto …).

    Im Mailtext eingebettete Bilder (Logos/Icons aus Signatur und Textkörper)
    werden aussortiert – der Nutzer will nur die E-Mail als PDF und die echten
    Anhänge. Nicht-Bild-Anhänge bleiben IMMER erhalten.
    """
    basis = os.path.basename(name or "")
    ext = os.path.splitext(basis)[1].lower()
    ist_bild = (ctype or "").lower().startswith("image/") or ext in _BILD_ENDUNGEN
    if not ist_bild:
        return True   # Dokumente (PDF/Word/Excel/…) immer behalten
    if inline:
        return False  # im Text eingebettetes Bild (Content-ID) -> Logo/Signatur
    if re.match(r"(?i)^(image\d{1,4}|oledata|emf\d*|clip_image\d*|~wrd\d*|att\d+)\.", basis):
        return False  # typische Auto-Namen eingebetteter Bilder
    groesse = len(data or b"")
    return not (groesse and groesse < 40 * 1024)   # winzige Bilder = Icons/Logos


def extrahiere(pfad):
    """Zerlegt eine E-Mail; liefert (erzeugte Dateipfade, Kontext-Text, Betreff).

    Der Kontext-Text (Betreff + Auszug) dient der Projekt-Zuordnung, damit alle
    aus EINER Mail erzeugten Dateien (Mailtext-PDF + Anhänge) demselben Projekt
    zugeordnet werden können. Der Betreff dient als Default-Bezeichnung für die
    Mailtext-PDF. Kleine Signatur-Logos werden übersprungen.
    """
    ordner = os.path.dirname(os.path.abspath(pfad))
    stem = os.path.splitext(os.path.basename(pfad))[0]
    endung = os.path.splitext(pfad)[1].lower()

    if endung == ".eml":
        kopf, body, anhaenge = _eml_lesen(pfad)
    elif endung == ".msg":
        kopf, body, anhaenge = _msg_lesen(pfad)
    else:
        return [], "", ""

    erzeugt = []
    pdf_name = _freier_name(ordner, _sicherer_dateiname(stem) + "_Mailtext", ".pdf")
    _text_zu_pdf(kopf, body, os.path.join(ordner, pdf_name))
    erzeugt.append(os.path.join(ordner, pdf_name))

    uebersprungen = 0
    for name, data, ctype, inline in anhaenge:
        if not data:
            continue
        if not _ist_echter_anhang(name, data, ctype, inline):
            uebersprungen += 1
            continue   # eingebettete Logos/Signatur-Bilder nicht extrahieren
        sicher = _sicherer_dateiname(name)
        st, ext = os.path.splitext(sicher)
        ziel = _freier_name(ordner, st or "Anhang", ext)
        with open(os.path.join(ordner, ziel), "wb") as f:
            f.write(data)
        erzeugt.append(os.path.join(ordner, ziel))

    kontext = (kopf.get("Betreff", "") + " " + (body or "")[:3000]).strip()
    return erzeugt, kontext, kopf.get("Betreff", "")


