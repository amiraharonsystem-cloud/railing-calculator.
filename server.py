import os
import io
import requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF
from arabic_reshaper import reshape
from bidi.algorithm import get_display

app = Flask(__name__)
CORS(app)

def fix_heb(text):
    if not text: return ""
    # סידור העברית מימין לשמאל ובצורה תקנית
    reshaped_text = reshape(str(text))
    return get_display(reshaped_text)

@app.route('/api/schedule/<name>', methods=['GET', 'POST'])
def generate_pdf(name):
    try:
        pdf = FPDF()
        pdf.add_page()
        
        # הורדת פונט Arimo שתומך בעברית (Unicode)
        font_path = "/tmp/Arimo-Regular.ttf"
        if not os.path.exists(font_path):
            url = "https://github.com/google/fonts/raw/main/ofl/arimo/Arimo%5Bwght%5D.ttf"
            r = requests.get(url)
            with open(font_path, "wb") as f:
                f.write(r.content)
        
        pdf.add_font("Arimo", "", font_path)
        pdf.set_font("Arimo", size=14)

        # נתוני חישוב מה-URL (למשל ?fw=1693&l1=1.2)
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        l2 = float(request.args.get('l2', 1.0))
        f_max = max(fw, fw * 0.943)

        # כותרת ודוח
        pdf.cell(0, 10, fix_heb(f"דוח בדיקת מעקות - בודק: {name}"), ln=True, align='C')
        pdf.ln(10)
        
        # תוצאות החישובים (סעיפי התקן)
        pdf.cell(0, 10, fix_heb(f"עומס תכנוני קובע: {f_max:.2f} N/m"), ln=True, align='R')
        pdf.cell(0, 10, fix_heb(f"סעיף א (עומס על מאחז): {0.375 * l1 * f_max:.2f} N"), ln=True, align='R')
        pdf.cell(0, 10, fix_heb(f"סעיף ב (כוח בחיבור): {0.75 * l1 * f_max:.2f} N"), ln=True, align='R')
        pdf.cell(0, 10, fix_heb(f"סעיף ה (ניצב אמצעי): {f_max * (l1/2 + l2/2):.2f} N"), ln=True, align='R')

        # יצירת הקובץ לזיכרון
        pdf_output = io.BytesIO()
        pdf_output.write(pdf.output())
        pdf_output.seek(0)

        return send_file(
            pdf_output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Report_{name}.pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
