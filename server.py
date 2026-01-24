import os
import io
import requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF

app = Flask(__name__)
CORS(app)

# פונקציה להיפוך טקסט עבור עברית (BiDi)
def fix_heb(text):
    if not text: return ""
    return str(text)[::-1]

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data received"}), 400

        # יצירת אובייקט PDF
        pdf = FPDF()
        pdf.add_page()
        
        # הורדת פונט Arimo של גוגל ישירות לזיכרון
        font_url = "https://github.com/google/fonts/raw/main/ofl/arimo/Arimo%5Bwght%5D.ttf"
        font_data = requests.get(font_url).content
        font_path = "/tmp/HebrewFont.ttf" # תיקייה זמנית המותרת לכתיבה בשרתי Linux
        
        with open(font_path, "wb") as f:
            f.write(font_data)
            
        pdf.add_font("Hebrew", "", font_path)
        pdf.set_font("Hebrew", size=14)

        # חישובים
        fw = float(data.get('fw', 0))
        l1 = float(data.get('l1', 1.0))
        l2 = float(data.get('l2', 1.0))
        f_max = max(fw, fw * 0.943)

        # כתיבת תוכן (היפוך טקסט לעברית)
        pdf.cell(0, 10, fix_heb("דוח הנדסי - בדיקת מעקות"), ln=True, align='C')
        pdf.ln(10)
        pdf.cell(0, 10, f"{fix_heb('שם פרויקט')}: {fix_heb(data.get('project', 'כללי'))}", ln=True, align='R')
        pdf.cell(0, 10, f"{fix_heb('עומס תכנוני קובע')}: {f_max:.2f} N/m", ln=True, align='R')
        pdf.ln(5)
        
        # סעיפי התקן
        pdf.cell(0, 10, f"{0.375 * l1 * f_max:.2f} N : {fix_heb('סעיף א - עומס על מאחז')}", ln=True, align='R')
        pdf.cell(0, 10, f"{0.75 * l1 * f_max:.2f} N : {fix_heb('סעיף ב - כוח בחיבור')}", ln=True, align='R')
        pdf.cell(0, 10, f"{f_max * (l1/2 + l2/2):.2f} N : {fix_heb('סעיף ה - ניצב אמצעי')}", ln=True, align='R')

        # יצירת ה-PDF לתוך זיכרון (Buffer) במקום קובץ פיזי
        pdf_output = io.BytesIO()
        pdf_str = pdf.output(dest='S') # output כ-string/bytes
        pdf_output.write(pdf_str)
        pdf_output.seek(0)

        return send_file(
            pdf_output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='Railing_Report.pdf'
        )

    except Exception as e:
        print(f"Error: {str(e)}") # יודפס בלוגים של Render
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
