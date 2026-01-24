import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# הקישור הישיר לגליון שלך בפורמט CSV
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs/gviz/tq?tqx=out:csv&gid=1101827105"

@app.route('/')
def home():
    return """
    <div dir="rtl" style="font-family:Arial; text-align:center; padding:50px;">
        <h1>המערכת עודכנה ופעילה! 🚀</h1>
        <p>לחץ למטה כדי לראות את הפרויקט של אמיר אהרון:</p>
        <a href="/api/schedule/Amir?date=25/01/2026" style="font-size:20px; color:#1a4e8a;">צפה בדוח ל-25/01/2026</a>
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # קריאת הגליון
        df = pd.read_csv(SHEET_CSV_URL)
        df.columns = df.columns.str.strip()
        
        # סינון לפי בודק ותאריך
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return f"<h1 dir='rtl'>לא נמצאו נתונים לאמיר בתאריך {target_date}</h1>", 404

        r = row.iloc[0]
        
        # הכנת הנתונים
        data = {
            "date": target_date,
            "project": str(r.get('שם המזמין', 'חסר')),
            "address": str(r.get('כתובת האתר', 'חסר')),
            "order": str(r.get('מספר הזמנה', '0')),
            "inspector": "אמיר אהרון",
            "f_max": 1693.68, # ערך דוגמה - ישלים מהאקסל
            "sec_e": 1693.68
        }

        if request.args.get('download') == 'excel':
            return generate_excel(data)

        return render_template('index.html', **data)
    
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

def generate_excel(data):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path):
        return "Missing template.xlsx in GitHub", 404
        
    wb = load_workbook(template_path)
    ws = wb.active

    # הזרקה לתאים לפי המבנה שלך
    ws['B2'] = data['date']
    ws['E2'] = data['inspector']
    ws['B3'] = data['project']
    ws['E3'] = data['order']
    ws['B4'] = data['address']
    ws['F15'] = data['f_max']
    ws['F20'] = data['sec_e']

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name="Report.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
