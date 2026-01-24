import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF
import io

app = Flask(__name__)
CORS(app)

@app.route('/api/generate-pdf', methods=['POST', 'GET'])
def generate_pdf():
    try:
        # קבלת נתונים מה-Frontend
        if request.method == 'POST':
            data = request.get_json() or {}
        else:
            data = request.args
            
        # יצירת PDF בסיסי (ללא פונטים חיצוניים בשלב זה כדי למנוע קריסה)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        pdf.cell(200, 10, txt="Railing Test Report - IS 1142", ln=True, align='C')
        pdf.ln(10)
        
        # הוספת הנתונים מהטופס
        for key, value in data.items():
            pdf.cell(200, 10, txt=f"{key}: {value}", ln=True, align='L')
            
        # יצירת הקובץ בזיכרון
        pdf_output = io.BytesIO()
        pdf_str = pdf.output(dest='S')
        pdf_output.write(pdf_str)
        pdf_output.seek(0)

        return send_file(
            pdf_output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='Report.pdf'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
