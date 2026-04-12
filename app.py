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
                    # White rectangle to overwrite existing content
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
        (1, 204.0, 669.8, f"{d['tp_vorname']} {d['tp_nachname']}", 10),  # Chen Coaching Mitglied
        (1, 204.0, 638.1, d["datum"],                                10),  # Datum: oben
        (1, 278.0,  99.1, d["datum"],                                10),  # Datum, Ort unten (nach dem Unterstrich)
    ])

def make_07(d):
    return fill_overlay(PDFS["07"], [
        (1, 222.0, 59.0, d["datum"], 10),  # Datum*
    ])

def make_04(d):
    # MAK inline: nur die MAK-Nummer, ohne Namen
    tp_mak = d.get("tp_mak_nr", "")
    mak_tg = tp_mak  # nur Nummer, kein Name
    fields = [
        # Vorname Tippgeber (rect top=154.7 -> fill_y=670.7)
        (1,  66.0, 670.7, d["tp_vorname"],          10),
        (1, 309.0, 670.7, d["tp_nachname"],         10),
        # Straße: white box über Spaldingstraße, dann Tippgeber-Adresse
        (1,  66.0, 641.8, d["tp_strasse"],          10, 231, 16),
        (1, 309.0, 641.8, d["tp_plz_ort"],          10, 231, 16),
        # MAK-Nr inline (nur wenn vorhanden)
        (1, 433.0, 591.5, mak_tg,                    8),
        # Seite 2: Datum
        (2, 179.0, 444.0, d["datum"],               10),
    ]
    return fill_overlay(PDFS["04"], fields)

def make_05(d):
    # Seite 1:
    #   Makler (A): top=149.8 Vorname/Nachname/MAK (bereits gefüllt)
    #               top=178.8 Straße/PLZ (Spaldingstraße vorgedruckt -> white box)
    #   Tippgeber (B): top=259.8 Vorname/Nachname/Ausweis, top=288.8 Straße/PLZ
    # Seite 2: Provision, Makler Ort+Datum, Tippgeber Ort+Datum
    prov = d.get("provision_pct", "")
    fields = [
        # Tippgeber Vorname: label top=274.7 -> text_y=570.3
        (1,  66.0, 570.3, d["tp_vorname"],           10),
        (1, 255.0, 570.3, d["tp_nachname"],          10),
        (1, 450.0, 570.3, d["tp_ausweis"],           10),
        # Tippgeber Straße: label top=303.7 -> text_y=541.3
        (1,  66.0, 541.3, d["tp_strasse"],           10),
        (1, 309.0, 541.3, d["tp_plz_ort"],           10),
        # Seite 2: Provision % (rect top=278.9, fill_y=556.0)
        (2, 122.0, 556.0, prov,                      10),
        # Seite 2: Makler Datum only (HH already printed at top=709.3)
        (2, 179.0, 127.1, d["datum"],                10),
        # Seite 2: Tippgeber Ort+Datum (rect top=741.5, fill_y=86.4)
        (2,  55.0,  86.4, d["tp_ort"],               10),
        (2, 179.0,  86.4, d["datum"],                10),
    ]
    return fill_overlay(PDFS["05"], fields)

def make_06(d):
    # MAK-Nummer (rect top=417.3, h=21.6 -> fill_y=408.1)
    # Ort,Datum  (rect top=446.4, h=33.1 -> fill_y=367.5, HH schon gedruckt)
    return fill_overlay(PDFS["06"], [
        (1,  55.0, 408.1, FIXED["mak_nr"],    10),
        (1, 179.0, 367.5, d["datum"],         10),
    ])

def make_03(d):
    reader = PdfReader(str(PDFS["03"]))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        c = canvas.Canvas(packet, pagesize=(pw, ph))
        c.setFont("Helvetica", 10)
        if i == 0:
            # MAK-Konto (rect top=165.8 x0=238.3)
            mak = d.get("tp_mak_nr", "")
            if mak:
                c.drawString(240, 662.1, mak)
            # Name/Vorname: line at top=233.4 -> text_y=607.5
            c.drawString(63, 595.0, f"{d['tp_vorname']} {d['tp_nachname']}")
            # Straße: line at top=262.3 -> text_y=578.6
            c.drawString(63, 578.6, d["tp_strasse"])
            # PLZ Ort: line at top=262.3 x0=305 -> text_y=578.6
            c.drawString(307, 578.6, d["tp_plz_ort"])
            # Kontoinhaber: line at top=291.3 -> text_y=549.6
            c.drawString(63, 549.6, f"{d['tp_vorname']} {d['tp_nachname']}")
            # IBAN (rect top=291.7 x0=306.3 -> text_y=540.2)
            c.setFont("Helvetica", 9)
            c.drawString(308, 540.2, d["tp_iban"])
            # Geldinstitut: line at top=320.2 -> text_y=520.7
            c.setFont("Helvetica", 10)
            c.drawString(63, 520.7, d["tp_geldinstitut"])
            # BIC (rect top=320.3 x0=306.3 -> text_y=511.6)
            c.setFont("Helvetica", 9)
            c.drawString(308, 511.6, d["tp_bic"])
            # Steuernummer (rect top=320.3 x0=434.4 -> text_y=511.6)
            c.drawString(436, 511.6, d["tp_steuernummer"])
            c.setFont("Helvetica", 10)
            # Datum Vermittler (HH bereits gedruckt, nur Datum)
            c.drawString(179, 228.0, d["datum"])
            # Ort + Datum Tippgeber
            c.drawString(55,  148.4, d["tp_ort"])
            c.drawString(179, 148.4, d["datum"])
        elif i == 1:
            # Name (box top=106.5, fill_y=713.9)
            c.drawString(55, 713.9, f"{d['tp_vorname']} {d['tp_nachname']}")
            # Straße (box top=135.8, fill_y=684.3)
            c.drawString(55, 684.3, d["tp_strasse"])
            c.drawString(303, 684.3, d["tp_plz_ort"])
            # Ort+Datum unten (box top=629.8, fill_y=179.5)
            c.drawString(55,  179.5, d["tp_ort"])
            c.drawString(179, 179.5, d["datum"])
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
