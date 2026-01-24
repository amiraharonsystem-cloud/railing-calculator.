import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF
import requests

app = Flask(__name__)
CORS(app)

# פונקציה להיפוך טקסט עבור עברית ב-PDF
def fix_heb(text):
    if not text: return ""
    return str(text)[::-1]

def create_pdf_engine():
    pdf = FPDF()
    pdf.add_page()
    
    # הורדת פונט עברי Arimo כדי שהשרת ב-Render יזהה עברית
    font_path = "HebrewFont.ttf"
    if not os.path.exists(font_path):
        url = "https://github.com/google/fonts/raw/main/ofl/arimo/Arimo%5Bwght%5D.ttf"
        r = requests.get(url)
        with open(font_path, "wb") as f:
            f.write(r.content)
            
    pdf.add_font("Arimo", "", font_path)
    pdf.set_font("Arimo", size=14)
    return pdf

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.json
        pdf = create_pdf_engine()
        
        # לוגיקה של החישובים מהתמונות שלך
        fw = float(data.get('fw', 0))
        l1 = float(data.get('l1', 1.0))
        l2 = float(data.get('l2', 1.0))
        f_max = max(fw, fw * 0.943)
        
        # כותרת
        pdf.cell(0, 10, fix_heb("דוח בדיקת מעקות - ת\"י 1142"), ln=True, align='C')
        pdf.ln(10)
        
        # נתונים
        pdf.cell(0, 10, f"{fix_heb('פרויקט')}: {fix_heb(data.get('project', 'כללי'))}", ln=True, align='R')
        pdf.cell(0, 10, f"{fix_heb('עומס תכנוני')} (Fmax): {f_max:.2f} N/m", ln=True, align='R')
        pdf.ln(5)
        
        # סעיפי התקן
        pdf.cell(0, 10, f"{0.375 * l1 * f_max:.2f} N : {fix_heb('סעיף א - עומס על מאחז')}", ln=True, align='R')
        pdf.cell(0, 10, f"{0.75 * l1 * f_max:.2f} N : {fix_heb('סעיף ב - כוח בחיבור')}", ln=True, align='R')
        pdf.cell(0, 10, f"{f_max * (l1/2 + l2/2):.2f} N : {fix_heb('סעיף ה - ניצב אמצעי')}", ln=True, align='R')
        
        pdf_path = "output_report.pdf"
        pdf.output(pdf_path)
        
        return send_file(pdf_path, as_attachment=True)
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
