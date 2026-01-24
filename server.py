import os
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF

app = Flask(__name__)
CORS(app)

@app.route('/api/schedule/<inspector_name>', methods=['GET', 'POST'])
def handle_pdf(inspector_name):
    try:
        # יצירת PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=16)
        
        # כותרת ונתונים
        pdf.cell(200, 10, txt="Railing Inspection Report", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=f"Inspector: {inspector_name}", ln=True)
        
        # הוספת נתוני טופס אם קיימים
        data = request.args.to_dict()
        for key, value in data.items():
            pdf.cell(0, 10, txt=f"{key}: {value}", ln=True)

        # יצירה לזיכרון
        pdf_output = io.BytesIO()
        pdf_str = pdf.output(dest='S')
        pdf_output.write(pdf_str)
        pdf_output.seek(0)

        return send_file(
            pdf_output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Report_{inspector_name}.pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
