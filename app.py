from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import io
import json
import zipfile
import os
import urllib.request
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject
from reportlab.pdfgen import canvas

app = Flask(__name__)
CORS(app)

BASE_DIR = Path(__file__).parent

PDFS = {
    "01": BASE_DIR / "01 Chen Coaching Prinzipien.pdf",
    "02": BASE_DIR / "02 Vertriebspartnervertrag Chen Coaching.pdf",
    "03": BASE_DIR / "03 FF-Formular_Änderung der Bankverbindung auf Dritten_OKT_2025_DIGI.pdf",
    "04": BASE_DIR / "04 FF_Sicherungsabtretung_2025_DIGI_04 1 Kopie.pdf",
    "05": BASE_DIR / "05 FF_Tippgebervertrag_2025_DIGI 1.pdf",
    "06": BASE_DIR / "06 FF_Vertriebsregulatorik.pdf",
    "07": BASE_DIR / "07 FF- Depotanbindung fondsplattformen.pdf",
}

FIXED = {
    "firma":   "Chen Coaching FDL GmbH",
    "strasse": "Spaldingstraße 210",
    "plz_ort": "20097 Hamburg",
    "mak_nr":  "MAK194317",
}

# Korrekte Formel: rl_y = 841.89 - rect.TOP - 3
# (rect.TOP = pdfplumber-Koordinate von oben nach unten)

def fill_fillable(src_path, fields):
    reader = PdfReader(str(src_path))
    writer = PdfWriter()
    writer.append(reader)
    if "/AcroForm" in writer._root_object:
        try:
            writer.update_page_form_field_values(None, fields, auto_regenerate=False)
        except TypeError:
            for page in writer.pages:
                try:
                    writer.update_page_form_field_values(page, fields, auto_regenerate=False)
                except Exception:
                    pass
        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()

def fill_overlay(src_path, text_fields):
    reader = PdfReader(str(src_path))
    writer = PdfWriter()
    pages_map = {}
    for entry in text_fields:
        pnum = entry[0]
        pages_map.setdefault(pnum, []).append(entry[1:])
    for i, page in enumerate(reader.pages):
        pnum = i + 1
        if pnum in pages_map:
            packet = io.BytesIO()
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(pw, ph))
            for entry in pages_map[pnum]:
                x, y, text, fsize = entry[0], entry[1], entry[2], entry[3]
                clear_w = entry[4] if len(entry) > 4 else 0
                clear_h = entry[5] if len(entry) > 5 else 0
                if clear_w and clear_h:
                    c.setFillColorRGB(1, 1, 1)
                    c.rect(x - 1, y - 3, clear_w, clear_h, fill=1, stroke=0)
                    c.setFillColorRGB(0, 0, 0)
                if text:
                    c.setFont("Helvetica", fsize)
                    c.drawString(x, y, text)
            c.save()
            packet.seek(0)
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
        writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()

def make_01(d):
    return fill_overlay(PDFS["01"], [
        (1, 204.0, 669.8, f"{d['tp_vorname']} {d['tp_nachname']}", 10),
        (1, 204.0, 638.1, d["datum"],                                10),
        (1, 278.0,  99.1, d["datum"],                                10),
    ])

def make_07(d):
    return fill_overlay(PDFS["07"], [
        (1, 222.0, 59.0, d["datum"], 10),
    ])

def make_04(d):
    # Direktes Canvas wie make_03 – zuverlässigste Methode
    # Formel: rl_y = 841.89 - rect.TOP - 3
    # Vorname/Name:  rect.top=147.0  → y=691.9
    # Straße/PLZ:    rect.top=176.7  → y=662.2  (weiße Box zuerst!)
    # MAK-Nr:        rect.top=228.5  → y=610.4
    # MAK §1.1:      rect.top=343.3  → y=495.6
    # P2 Datum:      rl_y = 461.8 - 8 = 434.3  (gemessen -8 relativ zu HH)
    tp_mak = d.get("tp_mak_nr", "")
    reader = PdfReader(str(PDFS["04"]))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        c = canvas.Canvas(packet, pagesize=(pw, ph))
        if i == 0:
            # Weiße Boxen über vorausgefüllte Felder (Straße/PLZ)
            c.setFillColorRGB(1, 1, 1)
            c.rect(57.4,  643.0, 242.1, 22.2, fill=1, stroke=0)
            c.rect(307.1, 643.0, 242.1, 22.2, fill=1, stroke=0)
            # Text schreiben
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 10)
            c.drawString(57.4,  691.9, d["tp_vorname"])
            c.drawString(307.1, 691.9, d["tp_nachname"])
            c.drawString(57.4,  662.2, d["tp_strasse"])
            c.drawString(307.1, 662.2, d["tp_plz_ort"])
            c.setFont("Helvetica", 9)
            c.drawString(435.4, 610.4, tp_mak)
            c.drawString(70.2,  495.6, tp_mak)
        elif i == 1:
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 10)
            c.drawString(177.0, 434.3, d["datum"])
        c.save()
        packet.seek(0)
        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()

def make_05(d):
    # Formel: rl_y = 841.89 - rect.TOP - 3
    # P1 Vorname/Nachname/Ausweis: rect.top=255.1 → y=583.8
    # P1 Straße/PLZ:               rect.top=284.9 → y=554.0
    # P2 Provision:                rect.top=274.8 → y=564.1
    # P2 Makler Datum:             rect.top=708.9 → y=130.0
    # P2 TG Ort/Datum:             rect.top=750.8 → y=88.1
    prov = d.get("provision_pct", "")
    fields = [
        (1,  57.4, 583.8, d["tp_vorname"],   10),
        (1, 251.3, 583.8, d["tp_nachname"],  10),
        (1, 452.3, 583.8, d["tp_ausweis"],   10),
        (1,  57.4, 554.0, d["tp_strasse"],   10),
        (1, 307.1, 554.0, d["tp_plz_ort"],   10),
        (2, 114.9, 564.1, prov,              10),
        (2, 173.7, 122.1, d["datum"],        10),   # HH rl_y=132.6, Datum -8 = 124 → gemessen 122
        (2,  46.0,  80.2, d["tp_ort"],       10),   # gemessen: Hamburg top=773.4 → rl_y=68.5, +12
        (2, 173.7,  80.2, d["datum"],        10),
    ]
    return fill_overlay(PDFS["05"], fields)

def make_06(d):
    return fill_overlay(PDFS["06"], [
        (1,  55.0, 408.1, FIXED["mak_nr"],    10),
        (1, 179.0, 367.5, d["datum"],         10),
    ])

def make_03(d):
    # Korrekte Formel: code_y = 841.89 - frame.TOP - 12
    # (Schrift rendert 7.9pt höher als angegeben, daher -12 statt -3)
    #
    # Page 1:
    #   MAK-Konto:   rect.top=158.4  → y=671.5
    #   Name/Vorn:   Linie top=228.0 → y=601.9
    #   Straße/PLZ:  Linie top=257.7 → y=572.2
    #   Kontoinhaber:Linie top=287.5 → y=542.4
    #   IBAN:        rect top=288.0  → y=541.9
    #   Geldinstitut:Linie top=317.3 → y=512.6
    #   BIC/Steuer:  rect top=317.36 → y=512.5
    #   Vermittler Datum: sign rect top=586.9 → y=243.0
    #   TG Ort/Datum: sign rect top=668.8 → y=161.1
    #
    # Page 2:
    #   Name:        rect top=97.3   → y=732.6
    #   Straße/PLZ:  rect top=127.5  → y=702.4
    #   TG sign:     rect top=635.9  → y=194.0
    reader = PdfReader(str(PDFS["03"]))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        c = canvas.Canvas(packet, pagesize=(pw, ph))
        c.setFont("Helvetica", 10)
        if i == 0:
            mak = d.get("tp_mak_nr", "")
            if mak:
                c.drawString(237.0, 673.6, mak)
            c.drawString(62.0, 604.0, f"{d['tp_vorname']} {d['tp_nachname']}")
            c.drawString(62.0, 574.3, d["tp_strasse"])
            c.drawString(306.0, 574.3, d["tp_plz_ort"])
            c.drawString(62.0, 544.5, f"{d['tp_vorname']} {d['tp_nachname']}")
            c.setFont("Helvetica", 9)
            c.drawString(307.0, 544.0, d["tp_iban"])
            c.setFont("Helvetica", 10)
            c.drawString(62.0, 514.7, d["tp_geldinstitut"])
            c.setFont("Helvetica", 9)
            c.drawString(307.0, 514.6, d["tp_bic"])
            c.drawString(438.0, 514.6, d["tp_steuernummer"])
            c.setFont("Helvetica", 10)
            c.drawString(174.0, 245.0, d["datum"])
            c.drawString(46.0,  163.1, d["tp_ort"])
            c.drawString(174.0, 163.1, d["datum"])
        elif i == 1:
            c.drawString(46.0,  734.6, f"{d['tp_vorname']} {d['tp_nachname']}")
            c.drawString(46.0,  704.5, d["tp_strasse"])
            c.drawString(301.3, 704.5, d["tp_plz_ort"])
            c.drawString(46.0,  196.1, d["tp_ort"])
            c.drawString(174.0, 196.1, d["datum"])
        c.save()
        packet.seek(0)
        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()

def make_02(d):
    page_h = 841.9
    def pdf_y(bottom): return page_h - bottom + 2
    return fill_overlay(PDFS["02"], [
        (1,  70.8,  pdf_y(197.9), d["tp_nachname"], 10),
        (1,  318.6, pdf_y(197.9), d["tp_vorname"],  10),
        (1,  70.8,  pdf_y(241.8), d["tp_strasse"],  10),
        (1,  318.6, pdf_y(241.8), d["tp_plz_ort"],  10),
        (10, 70.8,  pdf_y(305.4), f"{d['ort_datum']}, {d['datum']}", 10),
        (10, 70.8,  pdf_y(481.1), f"{d['tp_vorname']} {d['tp_nachname']}", 10),
    ])

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/lookup-iban", methods=["POST"])
def lookup_iban():
    try:
        data = request.get_json()
        iban = data.get("iban", "").replace(" ", "").upper()
        if not iban or len(iban) < 15:
            return jsonify({"valid": False})
        url = f"https://openiban.com/validate/{iban}?getBIC=true&validateBankCode=true"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            result = json.loads(resp.read().decode())
        if result.get("valid"):
            bank = result.get("bankData", {})
            return jsonify({"valid": True, "bic": bank.get("bic", ""), "name": bank.get("name", "")})
        else:
            return jsonify({"valid": False})
    except Exception as e:
        return jsonify({"valid": False, "error": str(e)})

@app.route("/generate", methods=["POST"])
def generate():
    try:
        d = request.get_json()
        if not d:
            return jsonify({"error": "Keine Daten empfangen"}), 400
        required = ["tp_vorname", "tp_nachname", "tp_strasse", "tp_plz_ort",
                    "tp_ausweis", "tp_iban", "tp_bic",
                    "tp_geldinstitut", "tp_steuernummer", "datum", "ort_datum"]
        missing = [f for f in required if not d.get(f)]
        if missing:
            return jsonify({"error": f"Fehlende Felder: {', '.join(missing)}"}), 400
        files = {
            "01_Chen_Coaching_Prinzipien.pdf": make_01(d),
            "02_Vertriebspartnervertrag.pdf":  make_02(d),
            "03_Bankverbindung.pdf":           make_03(d),
            "04_Sicherungsabtretung.pdf":      make_04(d),
            "05_Tippgebervertrag.pdf":         make_05(d),
            "06_Vertriebsregulatorik.pdf":     make_06(d),
            "07_Depotanbindung.pdf":           make_07(d),
        }
        zip_buf = io.BytesIO()
        name = f"{d['tp_nachname']}_{d['tp_vorname']}"
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, pdf_bytes in files.items():
                zf.writestr(f"{name}/{filename}", pdf_bytes)
        zip_buf.seek(0)
        return send_file(zip_buf, mimetype="application/zip", as_attachment=True,
                         download_name=f"{name}_Vertragsunterlagen.zip")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
