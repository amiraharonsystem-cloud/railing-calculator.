import os
import io
from flask import Flask, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app) # מאפשר לדפדפן לשלוח נתונים מהגיליון לשרת

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    try:
        template_path = 'template.xlsx'
        wb = load_workbook(template_path)
        ws = wb.active
        
        # הזרקת הנתונים שהגיעו מהדפדפן
        ws['B2'] = data.get('date', '')
        ws['E2'] = "אמיר אהרון"
        ws['B3'] = data.get('project', '')
        ws['E3'] = data.get('order', '')
        ws['B4'] = data.get('address', '')
        ws['F15'] = 1693.68
        ws['F20'] = 1693.68

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Report_{data.get('date')}.xlsx")
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
