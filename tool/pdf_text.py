"""
Pure-Python PDF-Textextraktion (ohne externe Abhaengigkeiten).

Dekodiert FlateDecode-Streams und nutzt die ToUnicode-CMaps der Subset-Fonts,
sodass auch Dokumente mit eigener Font-Kodierung lesbar werden.

Grenze: Rein bild-basierte Scans ohne Textebene koennen so nicht gelesen
werden - dafuer ist der API-Modus (mit OCR durch Claude) gedacht.
"""

import re
import zlib


def _parse_tounicode(stream):
    cmap = {}
    txt = stream.decode("latin-1", "replace")
    for blk in re.findall(r"beginbfchar(.*?)endbfchar", txt, re.DOTALL):
        for src, dst in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", blk):
            code = int(src, 16)
            u = "".join(chr(int(dst[i:i + 4], 16)) for i in range(0, len(dst), 4))
            cmap[code] = u
    for blk in re.findall(r"beginbfrange(.*?)endbfrange", txt, re.DOTALL):
        for lo, hi, dst in re.findall(
                r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", blk):
            a, b, base = int(lo, 16), int(hi, 16), int(dst, 16)
            for k, code in enumerate(range(a, b + 1)):
                cmap[code] = chr(base + k)
    return cmap


def _get_stream(body):
    m = re.search(rb"stream\r?\n(.*?)\r?\nendstream", body, re.DOTALL)
    if not m:
        return None
    raw = m.group(1)
    if b"/FlateDecode" in body:
        for candidate in (raw, raw[:-1], raw[:-2]):
            try:
                return zlib.decompress(candidate)
            except Exception:
                continue
        return raw
    return raw


def _read_string(s, j, L):
    depth, j = 1, j + 1
    cur = bytearray()
    while j < L and depth > 0:
        ch = s[j:j + 1]
        if ch == b"\\":
            nxt = s[j + 1:j + 2]
            mp = {b"n": 10, b"r": 13, b"t": 9, b"b": 8, b"f": 12,
                  b"(": 40, b")": 41, b"\\": 92}
            if nxt in mp:
                cur.append(mp[nxt]); j += 2; continue
            mo = re.match(rb"[0-7]{1,3}", s[j + 1:j + 4])
            if mo:
                cur.append(int(mo.group(), 8) & 0xFF); j += 1 + len(mo.group()); continue
            cur += nxt; j += 2; continue
        if ch == b"(":
            depth += 1; cur += ch; j += 1; continue
        if ch == b")":
            depth -= 1
            if depth == 0:
                j += 1; break
            cur += ch; j += 1; continue
        cur += ch; j += 1
    return bytes(cur), j


def _decode_content(content, fontmap):
    out, cur = [], None
    i, L, s = 0, len(content), content
    while i < L:
        ch = s[i:i + 1]
        if ch == b"/":
            m = re.match(rb"/[A-Za-z0-9_.+\-]+", s[i:i + 60])
            if m:
                name = m.group(0)
                i += len(name)
                if re.match(rb"\s+[-\d.]+\s+Tf", s[i:i + 40]):
                    cur = name
                continue
        if ch == b"(":
            b, i = _read_string(s, i, L)
            cm = fontmap.get(cur)
            out.append("".join(cm.get(x, "") for x in b) if cm
                       else b.decode("latin-1", "replace"))
            continue
        if ch == b"[":
            j = i + 1
            while j < L and s[j:j + 1] != b"]":
                if s[j:j + 1] == b"(":
                    b, j = _read_string(s, j, L)
                    cm = fontmap.get(cur)
                    out.append("".join(cm.get(x, "") for x in b) if cm
                               else b.decode("latin-1", "replace"))
                else:
                    j += 1
            i = j + 1
            continue
        if ch in (b"\n", b"\r") or s[i:i + 2] in (b"Td", b"TD", b"T*"):
            out.append("\n"); i += 1; continue
        i += 1
    return "".join(out)


def _resources_font_map(body, objs, font_to_cmap):
    rm = re.search(rb"/Resources\s+(\d+)\s+\d+\s+R", body)
    res = objs.get(int(rm.group(1)), b"") if rm else body
    fm = re.search(rb"/Font\s*<<(.*?)>>", res, re.DOTALL)
    block = fm.group(1) if fm else res
    fonts = {}
    for name, ref in re.findall(rb"/([A-Za-z0-9_.+\-]+)\s+(\d+)\s+\d+\s+R", block):
        fonts[b"/" + name] = font_to_cmap.get(int(ref))
    return fonts


def extract_pdf_text(path, max_chars=20000):
    """Liefert den extrahierten Text eines PDFs (leer, wenn nichts lesbar)."""
    try:
        data = open(path, "rb").read()
    except OSError:
        return ""

    objs = {}
    for m in re.finditer(rb"(\d+)\s+\d+\s+obj(.*?)endobj", data, re.DOTALL):
        objs[int(m.group(1))] = m.group(2)

    cmaps = {}
    for num, body in objs.items():
        st = _get_stream(body)
        if st and (b"beginbfchar" in st or b"beginbfrange" in st):
            cmaps[num] = _parse_tounicode(st)

    font_to_cmap = {}
    for num, body in objs.items():
        m = re.search(rb"/ToUnicode\s+(\d+)\s+\d+\s+R", body)
        if m and int(m.group(1)) in cmaps:
            font_to_cmap[num] = cmaps[int(m.group(1))]

    seiten = []
    for num, body in objs.items():
        if not re.search(rb"/Type\s*/Page\b", body):
            continue
        cont = b""
        cm = re.search(rb"/Contents\s+(\d+)\s+\d+\s+R", body)
        if cm:
            cont = _get_stream(objs.get(int(cm.group(1)), b"")) or b""
        else:
            arr = re.search(rb"/Contents\s*\[(.*?)\]", body, re.DOTALL)
            if arr:
                for r in re.findall(rb"(\d+)\s+\d+\s+R", arr.group(1)):
                    cont += (_get_stream(objs.get(int(r), b"")) or b"") + b"\n"
        if not cont:
            continue
        fontmap = _resources_font_map(body, objs, font_to_cmap)
        seiten.append((num, _decode_content(cont, fontmap)))

    seiten.sort()
    text = "\n".join(t for _, t in seiten)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text[:max_chars].strip()
