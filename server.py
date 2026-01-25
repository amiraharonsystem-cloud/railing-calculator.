import os
import io
from flask import Flask, request, send_file
from flask_cors import CORS
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
CORS(app)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        template_path = 'template.xlsx'
        
        if os.path.exists(template_path):
            wb = load_workbook(template_path)
        else:
            # אם הקובץ חסר, יוצר קובץ זמני כדי שלא תהיה שגיאה
            wb = Workbook()
            
        ws = wb.active
        ws['B2'] = data.get('date', '')
        ws['E2'] = "אמיר אהרון"
        ws['B3'] = data.get('project', '')
        ws['E3'] = data.get('order', '')
        ws['B4'] = data.get('address', '')

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name="Report.xlsx")
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
