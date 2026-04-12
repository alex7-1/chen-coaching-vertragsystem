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
    for (pnum, x, y, text, fsize) in text_fields:
        pages_map.setdefault(pnum, []).append((x, y, text, fsize))
    for i, page in enumerate(reader.pages):
        pnum = i + 1
        if pnum in pages_map:
            packet = io.BytesIO()
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(pw, ph))
            for (x, y, text, fsize) in pages_map[pnum]:
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
    mak_tg = f"{d.get('tp_mak_nr','')} / {d['tp_vorname']} {d['tp_nachname']}".strip("/ ")
    return fill_overlay(PDFS["04"], [
        # Page 1: Firma (schon vorausgefüllt, kein Overlay nötig)
        # Vorname Tippgeber
        (1,  66.0, 672.5, d["tp_vorname"],   10),
        # Nachname Tippgeber
        (1, 309.0, 672.5, d["tp_nachname"],  10),
        # MAK-Nr Tippgeber
        (1, 450.0, 672.5, d.get("tp_mak_nr",""), 10),
        # Straße Tippgeber
        (1,  66.0, 643.5, d["tp_strasse"],   10),
        # PLZ Ort Tippgeber
        (1, 309.0, 643.5, d["tp_plz_ort"],   10),
        # MAK-Nr + Name (inline Textfeld)
        (1, 435.0, 594.0, mak_tg,             8),
        # Page 2: Ort + Datum Sicherungsgeber
        (2,  57.0, 447.0, d["ort_datum"],    10),
        (2, 179.0, 447.0, d["datum"],        10),
    ])

def make_05(d):
    return fill_overlay(PDFS["05"], [
        # Page 1 - Tippgeber Daten
        # Vorname Tippgeber
        (1,  66.0, 706.0, d["tp_vorname"],        10),
        # Nachname Tippgeber
        (1, 268.0, 706.0, d["tp_nachname"],       10),
        # Personalausweisnummer
        (1, 470.0, 706.0, d["tp_ausweis"],        10),
        # Straße Tippgeber
        (1,  66.0, 677.0, d["tp_strasse"],        10),
        # PLZ Ort Tippgeber
        (1, 309.0, 677.0, d["tp_plz_ort"],        10),
        # Provisionssatz
        (1, 118.0, 536.5, d.get("provision_pct",""), 10),
        # Page 2 - Ort/Datum Makler (oben)
        (2,  55.0, 111.5, d["ort_datum"],         10),
        (2, 179.0, 111.5, d["datum"],             10),
        # Ort/Datum Tippgeber (unten)
        (2,  55.0,  70.5, d["tp_ort"],            10),
        (2, 179.0,  70.5, d["datum"],             10),
    ])

def make_06(d):
    return fill_overlay(PDFS["06"], [
        # MAK-Nummer (schon als HH vorausgefüllt - nur Datum fehlt)
        (1,  55.0, 406.0, FIXED["mak_nr"],              10),
        (1,  55.0, 371.5, f"{d['ort_datum']}, {d['datum']}", 10),
    ])

def make_03(d):
    pdf_bytes = fill_fillable(PDFS["03"], {
        "MAK-Konto": d.get("tp_mak_nr", ""),
        "Name Vorname  Firmenname": f"{d['tp_vorname']} {d['tp_nachname']}",
        "Straße Hausnummer": d["tp_strasse"],
        "PLZ Ort": d["tp_plz_ort"],
        "Kontoinhaber": f"{d['tp_vorname']} {d['tp_nachname']}",
        "IBAN": d["tp_iban"], "BIC": d["tp_bic"],
        "Geldinstitut": d["tp_geldinstitut"],
        "Steuernummer": d["tp_steuernummer"],
        "Ort":   d["ort_datum"],
        "Datum": d["datum"],
        "Ort_2":  d["tp_ort"],
        "Datum_2": d["datum"],
        "Straße Hausnummer Tippgeber": d["tp_strasse"],
        "PLZ Ort Tippgeber": d["tp_plz_ort"],
        "Datum_3": d["datum"],
        "Ort_3":   d["tp_ort"],
        "Gruppe1":  "/Auswahl2",
    })
    tmp_path = io.BytesIO(pdf_bytes)
    reader = PdfReader(tmp_path)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == 1:
            packet = io.BytesIO()
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(pw, ph))
            c.setFont("Helvetica", 10)
            # Name Vorname (oben, volle Breite)
            c.drawString(55, 717.5, f"{d['tp_vorname']} {d['tp_nachname']}")
            # Straße Hausnummer
            c.drawString(55, 687.5, d["tp_strasse"])
            # PLZ Ort
            c.drawString(303, 687.5, d["tp_plz_ort"])
            # Ort (Tippgeber-Ort, unten links)
            c.drawString(55, 645.0, d["tp_ort"])
            # Datum (unten mitte)
            c.drawString(179, 645.0, d["datum"])
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
