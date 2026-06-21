
import os
from datetime import datetime
import pandas as pd
import pdfkit
from flask import Flask, render_template, request, send_file, abort, redirect, url_for, send_from_directory

app = Flask(__name__, template_folder='.')
app.secret_key = "super_secret_session_key"

# Global variables
EXCEL_PATH = None
PHOTOS_FOLDER = None

# wkhtmltopdf paths
POSSIBLE_WKHTML_PATHS = [
    '/usr/bin/wkhtmltopdf',
    r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',
    r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
]

WKHTMLTOPDF_PATH = next((p for p in POSSIBLE_WKHTML_PATHS if os.path.exists(p)), POSSIBLE_WKHTML_PATHS[0])

# ------------------------------
# DATA FUNCTIONS
# ------------------------------
def load_data():
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        return pd.DataFrame()

    df = pd.read_excel(EXCEL_PATH)

    if not df.empty:
        if 'AssetBarcode' in df.columns:
            df['AssetBarcode'] = df['AssetBarcode'].astype(str).str.strip()
        if 'PIN' in df.columns:
            df['PIN'] = df['PIN'].astype(str).str.strip()

    return df


def calculate_age(purchase_date):
    if pd.isna(purchase_date):
        return None
    try:
        return (datetime.now() - pd.to_datetime(purchase_date)).days / 365.25
    except:
        return None


def get_recommendation(condition, system_type, purchase_date):
    cond = str(condition).lower()
    age = calculate_age(purchase_date)

    if "poor" in cond:
        return "IMMEDIATE REPAIR / REPLACE"
    if "fair" in cond:
        return "PLAN MAINTENANCE"
    if "good" in cond:
        return "ROUTINE CHECK"
    return "INSPECT MANUALLY"


def get_photo_filename(barcode):
    if not PHOTOS_FOLDER or not os.path.exists(PHOTOS_FOLDER):
        return None

    for ext in ['.jpg', '.png', '.jpeg']:
        file = f"{barcode}{ext}"
        if os.path.exists(os.path.join(PHOTOS_FOLDER, file)):
            return file
    return None


# ------------------------------
# ROUTES
# ------------------------------
@app.route('/', methods=['GET', 'POST'])
def index():
    global EXCEL_PATH, PHOTOS_FOLDER

    if request.method == 'POST':

        # Upload Excel
        if 'excel_file' in request.files:
            file = request.files['excel_file']
            if file.filename:
                EXCEL_PATH = os.path.join(os.getcwd(), file.filename)
                file.save(EXCEL_PATH)

        # ✅ User-defined folder
        folder = request.form.get('photos_folder')
        if folder and os.path.exists(folder):
            PHOTOS_FOLDER = folder

        return redirect(url_for('index'))

    df = load_data()

    search_barcode = request.args.get('barcode', '').strip()

    selected_asset = None
    photo_url = None

    if not df.empty and search_barcode:
        result = df[df['AssetBarcode'] == search_barcode]
        if not result.empty:
            selected_asset = result.iloc[0].to_dict()

            photo_file = get_photo_filename(search_barcode)
            if photo_file:
                photo_url = url_for('serve_photo', filename=photo_file)

    return render_template(
        'index.html',
        selected_asset=selected_asset,
        photo_url=photo_url
    )


# ✅ Serve images dynamically
@app.route('/photo/<filename>')
def serve_photo(filename):
    if not PHOTOS_FOLDER:
        abort(404)
    return send_from_directory(PHOTOS_FOLDER, filename)


# ------------------------------
# PDF GENERATION (unchanged)
# ------------------------------
@app.route('/generate-pdf')
def generate_pdf():
    df = load_data()
    if df.empty:
        return "No data", 400

    html = "<h1>Report</h1>"

    path = "report.pdf"
    config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

    pdfkit.from_string(html, path, configuration=config)

    return send_file(path, as_attachment=True)


# ------------------------------
if __name__ == "__main__":
    app.run(debug=True)
