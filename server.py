import os
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF

app = Flask(__name__)
CORS(app)

def fix_heb(text):
    return str(text)[::-1]

@app.route('/api/schedule/<name>', methods=['GET', 'POST'])
def generate_report(name):
    try:
        # יצירת PDF
        pdf = FPDF()
        pdf.add_page()
        # שימוש בפונט בסיסי (בשרת מרוחק עדיף Arial או Courier ללא התקנה מורכבת)
        pdf.set_font("Helvetica", size=12)
        
        # נתונים לדוגמה מהחישובים שלך (ניתן להעביר אותם ב-URL)
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        l2 = float(request.args.get('l2', 1.0))
        f_max = max(fw, fw * 0.943)

        pdf.cell(200, 10, txt=f"Engineering Report for: {name}", ln=True, align='C')
        pdf.ln(10)
        
        # חישובים לפי התקן
        pdf.cell(0, 10, txt=f"Design Load (Fmax): {f_max:.2f} N/m", ln=True)
        pdf.cell(0, 10, txt=f"Section A (Handrail): {0.375 * l1 * f_max:.2f} N", ln=True)
        pdf.cell(0, 10, txt=f"Section B (Connection): {0.75 * l1 * f_max:.2f} N", ln=True)
        pdf.cell(0, 10, txt=f"Section E (Middle Post): {f_max * (l1/2 + l2/2):.2f} N", ln=True)

        pdf_output = io.BytesIO()
        pdf_bytes = pdf.output(dest='S').encode('latin-1', 'ignore')
        pdf_output.write(pdf_bytes)
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
