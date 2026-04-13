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
    """
    PDF 04 Sicherungsabtretung
    Koordinaten ermittelt via pdfplumber (top -> reportlab y = 841.9 - top)
    
    Seite 1:
      - Firma:    label top=132.4  -> y≈694
      - Vorname:  label top=162.1  -> y≈665
      - Name:     x=309.7, gleiche Zeile
      - Straße:   label top=191.8  -> y≈635
      - PLZ,Ort:  x=309.7, gleiche Zeile
      - MAK-Nr (Header-Box): x=436, label top=243.6 -> y≈594
      - MAK inline §1.1: nach 'MAK-Nr. / Name des Tippgebers', top=349.9 -> y≈488
    Seite 2:
      - Datum: x=173, label top=411.4 -> y≈433
    """
    tp_mak = d.get("tp_mak_nr", "")
    fields = [
        # Seite 1: Tippgeber-Daten in den Feldern des Sicherungsgebers
        (1,  60.0, 694.0, d["tp_vorname"],          10),
        (1, 309.7, 694.0, d["tp_nachname"],          10),
        (1,  60.0, 635.0, d["tp_strasse"],           10),
        (1, 309.7, 635.0, d["tp_plz_ort"],           10),
        # MAK-Nr. Feld oben rechts (Header-Box)
        (1, 436.0, 594.0, tp_mak,                     8),
        # MAK inline in §1.1 nach "MAK-Nr. / Name des Tippgebers"
        (1, 165.0, 488.0, tp_mak,                     8),
        # Seite 2: Datum
        (2, 173.0, 433.0, d["datum"],                10),
    ]
    return fill_overlay(PDFS["04"], fields)

def make_05(d):
    """
    PDF 05 Tippgebervertrag
    Koordinaten ermittelt via pdfplumber

    Seite 1 - Tippgeber (Block B):
      - Vorname:  label top=270.3 -> y≈560
      - Nachname: x=253.6, gleiche Zeile
      - Ausweis:  x=454.9, gleiche Zeile
      - Straße:   label top=300.1 -> y≈530
      - PLZ,Ort:  x=309.7, gleiche Zeile

    Seite 2:
      - Provision %: vor dem '%'-Zeichen, x=116, top=277.3 -> y≈567
      - Makler Datum: x=176, top=735.4 -> y≈109  (Ort "HH" bereits gedruckt)
      - TG Ort:       x=48.6, top=777.3 -> y≈67
      - TG Datum:     x=176,  top=777.3 -> y≈67
    """
    prov = d.get("provision_pct", "")
    fields = [
        # Seite 1: Tippgeber Vorname/Nachname/Ausweis
        (1,  60.0, 560.0, d["tp_vorname"],           10),
        (1, 253.6, 560.0, d["tp_nachname"],          10),
        (1, 454.9, 560.0, d["tp_ausweis"],           10),
        # Seite 1: Tippgeber Straße / PLZ Ort
        (1,  60.0, 530.0, d["tp_strasse"],           10),
        (1, 309.7, 530.0, d["tp_plz_ort"],           10),
        # Seite 2: Provision %
        (2, 116.0, 567.0, prov,                      10),
        # Seite 2: Makler Datum (Ort "HH" bereits aufgedruckt)
        (2, 176.0, 109.0, d["datum"],                10),
        # Seite 2: Tippgeber Ort + Datum
        (2,  48.6,  67.0, d["tp_ort"],               10),
        (2, 176.0,  67.0, d["datum"],                10),
    ]
    return fill_overlay(PDFS["05"], fields)

def make_06(d):
    return fill_overlay(PDFS["06"], [
        (1,  55.0, 408.1, FIXED["mak_nr"],    10),
        (1, 179.0, 367.5, d["datum"],         10),
    ])

def make_03(d):
    """
    PDF 03 Änderung der Bankverbindung auf Dritten
    Koordinaten ermittelt via pdfplumber (top -> reportlab y = 841.9 - top)

    Seite 1:
      - MAK-Konto (A-Feld):   x=240, zwischen 'MAK-Konto*' (top=166.3) und 'hinterlegten' -> y≈684
      - Name/Vorname (B-Box): x=57,  label top=243.6 -> fill y≈584
      - Straße:               x=57,  label top=273.3 -> fill y≈554
      - PLZ,Ort:              x=308, gleiche Zeile   -> fill y≈554
      - Kontoinhaber:         x=57,  label top=303.1 -> fill y≈524
      - IBAN:                 x=308, label top=306.1 -> fill y≈524
      - Geldinstitut:         x=57,  label top=332.9 -> fill y≈495
      - BIC:                  x=308, label top=335.5 -> fill y≈495
      - Steuernummer:         x=438, label top=335.5 -> fill y≈495
      - Datum Vermittler:     x=179, label top=613.3 -> fill y≈237  (Ort "HH" gedruckt)
      - Ort Tippgeber:        x=48,  label top=695.2 -> fill y≈155
      - Datum Tippgeber:      x=176, gleiche Zeile   -> fill y≈155

    Seite 2:
      - Name/Vorname:         x=55,  label top=112.4 -> fill y≈714
      - Straße:               x=55,  label top=142.6 -> fill y≈685
      - PLZ,Ort:              x=304, gleiche Zeile   -> fill y≈685
      - Ort+Datum Tippgeber:  x=55/176, label top=662.2 -> fill y≈188
    """
    reader = PdfReader(str(PDFS["03"]))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        c = canvas.Canvas(packet, pagesize=(pw, ph))
        c.setFont("Helvetica", 10)
        if i == 0:
            # MAK-Konto (A-Block): inline nach "MAK-Konto*"
            mak = d.get("tp_mak_nr", "")
            if mak:
                c.drawString(240, 684.0, mak)
            # Name/Vorname des Tippgebers (B-Block)
            c.drawString(57, 584.0, f"{d['tp_vorname']} {d['tp_nachname']}")
            # Straße
            c.drawString(57, 554.0, d["tp_strasse"])
            # PLZ, Ort
            c.drawString(308, 554.0, d["tp_plz_ort"])
            # Kontoinhaber
            c.drawString(57, 524.0, f"{d['tp_vorname']} {d['tp_nachname']}")
            # IBAN
            c.setFont("Helvetica", 9)
            c.drawString(308, 524.0, d["tp_iban"])
            # Geldinstitut
            c.setFont("Helvetica", 10)
            c.drawString(57, 495.0, d["tp_geldinstitut"])
            # BIC
            c.setFont("Helvetica", 9)
            c.drawString(308, 495.0, d["tp_bic"])
            # Steuernummer
            c.drawString(438, 495.0, d["tp_steuernummer"])
            # Datum Vermittler (Ort "HH" bereits gedruckt)
            c.setFont("Helvetica", 10)
            c.drawString(179, 237.0, d["datum"])
            # Ort + Datum Tippgeber
            c.drawString(48,  155.0, d["tp_ort"])
            c.drawString(176, 155.0, d["datum"])
        elif i == 1:
            # Name/Vorname (top=112.4 -> y=714)
            c.drawString(55, 714.0, f"{d['tp_vorname']} {d['tp_nachname']}")
            # Straße (top=142.6 -> y=685)
            c.drawString(55, 685.0, d["tp_strasse"])
            c.drawString(304, 685.0, d["tp_plz_ort"])
            # Ort + Datum Tippgeber (top=662.2 -> y=188)
            c.drawString(55,  188.0, d["tp_ort"])
            c.drawString(176, 188.0, d["datum"])
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
