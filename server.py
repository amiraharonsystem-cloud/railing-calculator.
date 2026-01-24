import os
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF

app = Flask(__name__)
CORS(app)

@app.route('/api/schedule/<name>', methods=['GET', 'POST'])
def generate_pdf(name):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        # כותרת הדוח
        pdf.cell(200, 10, txt=f"Engineering Report - Inspector: {name}", ln=True, align='C')
        pdf.ln(10)
        
        # קבלת נתונים מה-URL (למשל ?fw=1693)
        fw = float(request.args.get('fw', 1693.68))
        pdf.cell(0, 10, txt=f"Design Load (Fmax): {fw:.2f} N/m", ln=True)
        
        # חישוב סעיף ה' (ניצב אמצעי)
        l1 = float(request.args.get('l1', 1.0))
        l2 = float(request.args.get('l2', 1.0))
        section_e = fw * (l1/2 + l2/2)
        pdf.cell(0, 10, txt=f"Section E (Middle Post Load): {section_e:.2f} N", ln=True)

        # יצירת הקובץ לזיכרון
        pdf_output = io.BytesIO()
        # קידוד latin-1 כדי למנוע קריסה על תווים מיוחדים בשלב זה
        pdf_content = pdf.output(dest='S').encode('latin-1', 'ignore')
        pdf_output.write(pdf_content)
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
    # שימוש בפורט ש-Render מצפה לו
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
