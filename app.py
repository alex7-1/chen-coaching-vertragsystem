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
    "02": BASE_DIR / "02 Vertriebspartnervertrag Chen Coaching.pdf",
    "03": BASE_DIR / "03 FF-Formular_Änderung der Bankverbindung auf Dritten_OKT_2025_DIGI.pdf",
    "04": BASE_DIR / "04 FF_Sicherungsabtretung_2025_DIGI_04 1 Kopie.pdf",
    "05": BASE_DIR / "05 FF_Tippgebervertrag_2025_DIGI 1.pdf",
    "06": BASE_DIR / "06 FF_Vertriebsregulatorik.pdf",
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
    try:
        writer.update_page_form_field_values(None, fields, auto_regenerate=False)
    except TypeError:
        for page in writer.pages:
            try:
                writer.update_page_form_field_values(page, fields, auto_regenerate=False)
            except Exception:
                pass
    # NeedAppearances: PDF-Viewer soll Felder neu rendern (wichtig für Comb-Felder wie IBAN/BIC)
    if "/AcroForm" in writer._root_object:
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

def make_04(d):
    return fill_fillable(PDFS["04"], {
        "Firmenname": FIXED["firma"],
        "Straße Hausnummer": FIXED["strasse"],
        "PLZ Ort": FIXED["plz_ort"],
        "MAK-Nummer": FIXED["mak_nr"],
        "MAK-Nummer und Name des Tippgebers": f"{d['tp_mak_nr']} / {d['tp_vorname']} {d['tp_nachname']}",
        "Ort_2": d["ort_datum"], "Datum_2": d["datum"],
        "Vorname": "", "Nachname": "",
    })

def make_05(d):
    return fill_fillable(PDFS["05"], {
        "Firmenname": FIXED["firma"],
        "Vorname Makler": "", "Nachname Makler": "",
        "Straße Hausnummer Makler": FIXED["strasse"],
        "PLZ Ort Makler": FIXED["plz_ort"],
        "Vorname Tippgeber": d["tp_vorname"],
        "Nachname Tippgeber": d["tp_nachname"],
        "Straße Hausnummer Tippgeber": d["tp_strasse"],
        "PLZ Ort Tippgeber": d["tp_plz_ort"],
        "Personalausweisnummer": d["tp_ausweis"],
        "Prozent der Abschlussprovision": d.get("provision_pct", ""),
        "Ort":  d["ort_datum"],   # Makler (Chen Coaching)
        "Datum": d["datum"],
        "Ort_2":  d["tp_ort"],    # Tippgeber
        "Datum_2": d["datum"],
    })

def make_06(d):
    return fill_fillable(PDFS["06"], {
        "MAK-Nummer": FIXED["mak_nr"],
        "Ort, Datum": f"{d['ort_datum']}, {d['datum']}",
    })

def make_03(d):
    # Schritt 1: AcroForm-Felder füllen
    pdf_bytes = fill_fillable(PDFS["03"], {
        "MAK-Konto": FIXED["mak_nr"],
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
    })
    # Schritt 2: Seite 2 – Name/Vorname Tippgeber per Overlay einfügen
    # pdfplumber top=112.4 → PDF y = 841.9 - 112.4 - 2 = 727
    page_h = 841.9
    tmp_path = io.BytesIO(pdf_bytes)
    reader = PdfReader(tmp_path)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == 1:  # Seite 2
            packet = io.BytesIO()
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(pw, ph))
            c.setFont("Helvetica", 10)
            c.drawString(48, 731.5, f"{d['tp_vorname']} {d['tp_nachname']}")
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
            "02_Vertriebspartnervertrag.pdf": make_02(d),
            "03_Bankverbindung.pdf":          make_03(d),
            "04_Sicherungsabtretung.pdf":     make_04(d),
            "05_Tippgebervertrag.pdf":        make_05(d),
            "06_Vertriebsregulatorik.pdf":    make_06(d),
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
