import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fpdf import FPDF
import io

app = Flask(__name__)
CORS(app)

def create_simple_pdf(data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Railing Report", ln=True, align='C')
    pdf.ln(10)
    
    # כתיבת הנתונים שהתקבלו
    for key, value in data.items():
        pdf.cell(0, 10, txt=f"{key}: {value}", ln=True)
        
    pdf_output = io.BytesIO()
    pdf_str = pdf.output(dest='S')
    pdf_output.write(pdf_str)
    pdf_output.seek(0)
    return pdf_output

# זה הנתיב שה-Frontend שלך מחפש לפי צילום המסך (image_15c593.jpg)
@app.route('/api/schedule/<name>', methods=['GET', 'POST'])
def schedule_pdf(name):
    try:
        # אם זו בקשת GET (כמו בדפדפן), ניקח נתונים מה-URL
        data = request.args.to_dict()
        data['Inspector'] = name
        
        pdf_file = create_simple_pdf(data)
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'Report_{name}.pdf'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
