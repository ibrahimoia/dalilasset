import os
import sys
from datetime import datetime
import pandas as pd
import pdfkit
from flask import Flask, render_template, request, send_file, abort, redirect, url_for

# --- Initialization & Configuration ---
app = Flask(__name__)
app.secret_key = "super_secret_session_key" 

# Track global runtime variables safely
EXCEL_PATH = None
PHOTOS_FOLDER = os.path.join(app.root_path, 'static', 'photos')

# 32-bit & 64-bit Compatibility: Broad fallback array for newer Windows OS environments

POSSIBLE_WKHTML_PATHS = [
    '/usr/bin/wkhtmltopdf',  # Default Linux path (Render/Ubuntu)
    r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',
    r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
]

#POSSIBLE_WKHTML_PATHS = [
    # r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',  # 32-bit app on 64-bit Windows
    #r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',        # Native 64-bit or Native 32-bit OS
#]
# Dynamic confirmation check
WKHTMLTOPDF_PATH = next((path for path in POSSIBLE_WKHTML_PATHS if os.path.exists(path)), POSSIBLE_WKHTML_PATHS[0])


def load_data():
    """Reads dataset from the dynamically selected active Excel workbook."""
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
    if pd.isna(purchase_date) or not purchase_date:
        return None
    try:
        if isinstance(purchase_date, str):
            p_date = datetime.strptime(purchase_date.split()[0], '%Y-%m-%d')
        elif hasattr(purchase_date, 'to_pydatetime'):
            p_date = purchase_date.to_pydatetime()
        else:
            p_date = purchase_date
        delta = datetime.now() - p_date
        return delta.days / 365.25
    except:
        return None

def get_recommendation(condition, system_type, purchase_date):
    cond = str(condition).strip().lower() if pd.notna(condition) else "unknown"
    age = calculate_age(purchase_date)
    
    if cond in ["poor", "سيء", "broken", "unstable"] or "poor" in cond:
        return "RECOMMENDATION: IMMEDIATE REPAIR / REPLACE. Asset requires swift intervention."
    if cond in ["fair", "متوسط"] or "satisfactory" in cond or "2" in cond:
        if age and age > 8:
            return "RECOMMENDATION: PLAN FOR REPLACEMENT. System lifecycle optimization suggested."
        return "RECOMMENDATION: SCHEDULE REPAIR/MAINTENANCE CYCLE."
    if cond in ["excellent", "new", "جديد", "good"] or "1" in cond:
        return "SCHEDULE: ROUTINE MONITORING SERVICE CYCLE."

    return "STATUS: Manual inspection verification recommended."

def get_photo_filename(barcode):
    extensions = ['.jpg', '.png', '.jpeg', '.JPG', '.PNG']
    for ext in extensions:
        filename = f"{barcode}{ext}"
        if os.path.exists(os.path.join(PHOTOS_FOLDER, filename)):
            return filename
    return None

# --- Web App Routes ---

@app.route('/', methods=['GET', 'POST'])
def index():
    global EXCEL_PATH
    
    # Handle File Upload directly from front end
    if request.method == 'POST':
        if 'excel_file' in request.files:
            file = request.files['excel_file']
            if file.filename != '' and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                upload_dir = os.getcwd()
                EXCEL_PATH = os.path.join(upload_dir, file.filename)
                file.save(EXCEL_PATH) 
                return redirect(url_for('index'))
                
    df = load_data()
    
    if df.empty:
        return render_template('index.html', needs_upload=True)

    search_barcode = request.args.get('barcode', '').strip()
    search_pin = request.args.get('pin', '').strip()
    
    selected_asset = None
    photo_url = None
    recommendation = ""

    # 1. Search Barcode Processing
    if not df.empty and search_barcode:
        result = df[df['AssetBarcode'] == search_barcode]
        if not result.empty:
            selected_asset = result.iloc[0].to_dict()
            recommendation = get_recommendation(
                selected_asset.get('AssetConditionsDescription'),
                selected_asset.get('sys'),
                selected_asset.get('date-purchase')
            )
            photo_file = get_photo_filename(search_barcode)
            if photo_file:
                photo_url = f"/static/photos/{photo_file}"

    # 2. Search PIN Processing (Pulls full collection for that PIN context)
    elif not df.empty and search_pin:
        pin_results = df[df['PIN'] == search_pin]
        if not pin_results.empty:
            # We target the first record found to display high level focus, or list it
            selected_asset = pin_results.iloc[0].to_dict()
            recommendation = get_recommendation(
                selected_asset.get('AssetConditionsDescription'),
                selected_asset.get('sys'),
                selected_asset.get('date-purchase')
            )
            photo_file = get_photo_filename(selected_asset.get('AssetBarcode'))
            if photo_file:
                photo_url = f"/static/photos/{photo_file}"

    all_assets = df.to_dict(orient='records') if not df.empty else []
    
    return render_template(
        'index.html', 
        needs_upload=False,
        assets=all_assets, 
        selected_asset=selected_asset,
        recommendation=recommendation,
        photo_url=photo_url,
        search_barcode=search_barcode,
        search_pin=search_pin
    )
@app.route('/guide')
def view_guide():
    """Renders the data preparation documentation guide."""
    return render_template('instructions.html')

@app.route('/generate-pdf', methods=['GET'])
def generate_pdf_route():
    mode = request.args.get('mode')
    key = request.args.get('key', '').strip()
    
    df = load_data()
    if df.empty:
        return "No data available. Please upload an Excel file first.", 400
        
    if mode == "Barcode":
        filtered_df = df[df['AssetBarcode'] == key]
        filename = f"Report_Barcode_{key.replace('/', '_')}.pdf"
    elif mode == "PIN":
        filtered_df = df[df['PIN'] == key]
        filename = f"Report_PIN_{key.replace('/', '_')}.pdf"
    else:
        return "Invalid Request Mode", 400

    if filtered_df.empty:
        return f"No records found for {mode}: {key}", 404

    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 30px; color: #333; }}
            h1 {{ color: #2C3E50; border-bottom: 2px solid #2196F3; padding-bottom: 10px; }}
            .report-meta {{ margin-bottom: 20px; font-style: italic; }}
            .asset-card {{ border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 25px; page-break-inside: avoid; }}
            .asset-title {{ font-size: 18px; font-weight: bold; color: #1565C0; margin-bottom: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 13px; }}
            th {{ background-color: #f5f5f5; width: 30%; }}
            .recommendation {{ margin-top: 10px; padding: 10px; background-color: #FFEBEE; color: #C62828; font-weight: bold; border-left: 4px solid #D32F2F; }}
            .photo-box {{ margin-top: 15px; max-width: 400px; max-height: 300px; border: 1px dashed #ccc; padding: 5px; }}
        </style>
    </head>
    <body>
        <h1>Asset Condition Assessment Report</h1>
        <div class="report-meta">Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Target Map: {key} ({mode} View)</div>
    """

    for _, row in filtered_df.iterrows():
        bc = row.get('AssetBarcode', 'N/A')
        cond = row.get('AssetConditionsDescription', 'N/A')
        sys_t = row.get('sys', 'N/A')
        p_dt = row.get('date-purchase', 'N/A')
        recom = get_recommendation(cond, sys_t, p_dt)
        
        photo_f = get_photo_filename(bc)
        if photo_f:
            abs_photo_path = os.path.join(PHOTOS_FOLDER, photo_f)
            img_html = f'<img class="photo-box" src="{abs_photo_path}">'
        else:
            img_html = '<p style="color:#777; font-style:italic;">No photo path associated with asset.</p>'

        html_content += f"""
        <div class="asset-card">
            <div class="asset-title">Asset Barcode: {bc}</div>
            <table>
                <tr><th>PIN ID</th><td>{row.get('PIN', 'N/A')}</td></tr>
                <tr><th>Description</th><td>{row.get('AssetDescription', 'N/A')}</td></tr>
                <tr><th>System Segment</th><td>{sys_t}</td></tr>
                <tr><th>Location Context</th><td>{row.get('building', 'N/A')} - {row.get('location', 'N/A')}</td></tr>
                <tr><th>Condition Recorded</th><td>{cond}</td></tr>
                <tr><th>Purchase Date</th><td>{p_dt}</td></tr>
                <tr><th>Inspector Comments</th><td>{row.get('comments', '')}</td></tr>
            </table>
            <div class="recommendation">{recom}</div>
            <div style="margin-top:10px;"><strong>Visual Documentation:</strong><br/>{img_html}</div>
        </div>
        """

    html_content += "</body></html>"

    pdf_output_path = os.path.join(os.getcwd(), 'temp_report.pdf')
    try:
        config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        pdfkit.from_string(html_content, pdf_output_path, configuration=config, options={"enable-local-file-access": ""})
        return send_file(pdf_output_path, as_attachment=True, download_name=filename)
    except Exception as e:
        return f"An internal exception occurred writing PDF compilation: {e}", 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
