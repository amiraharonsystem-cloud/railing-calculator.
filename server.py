import os
import io
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        df = pd.read_csv(SHEET_CSV_URL)
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return render_template('index.html', error="לא נמצאו נתונים")

        r = row.iloc[0]
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        f_max = max(fw, fw * 0.943)

        data = {
            "date": target_date,
            "project": r.get('שם המזמין', ''),
            "address": r.get('כתובת האתר', ''),
            "order": r.get('מספר הזמנה', ''),
            "inspector": "אמיר אהרון",
            "f_max": round(f_max, 2),
            "sec_e": round(f_max * l1, 2)
        }

        # אם המשתמש ביקש להוריד אקסל
        if request.args.get('download') == 'excel':
            return download_excel_filled(data)

        return render_template('index.html', **data)
    except Exception as e:
        return f"Error: {str(e)}"

def download_excel_filled(data):
    # טעינת התבנית המקורית מהתיקייה
    template_path = 'template.xlsx'
    wb = load_workbook(template_path)
    ws = wb.active # או ws = wb['שם הגיליון']

    # הזרקת הנתונים בדיוק לתאים המקוריים (דוגמה לפי המבנה שלך)
    ws['B2'] = data['date']
    ws['E2'] = data['inspector']
    ws['B3'] = data['project']
    ws['E3'] = data['order']
    ws['B4'] = data['address']
    ws['F15'] = data['f_max'] # דוגמה לתא חישוב
    ws['F20'] = data['sec_e'] # תא סעיף ה'

    # שמירה לזיכרון ושליחה למשתמש
    excel_out = io.BytesIO()
    wb.save(excel_out)
    excel_out.seek(0)
    
    return send_file(
        excel_out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"Report_{data['date'].replace('/','-')}.xlsx"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
