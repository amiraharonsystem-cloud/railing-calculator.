import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

SPREADSHEET_ID = "1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs"

@app.route('/')
def home():
    return """
    <div dir="rtl" style="font-family:Arial; text-align:center; padding:50px;">
        <h1>מערכת דוחות - אמיר אהרון</h1>
        <p>בחר תאריך מהלוח:</p>
        <input type="date" id="datePicker" style="padding:10px; font-size:18px;">
        <br><br>
        <button onclick="submitDate()" style="padding:10px 20px; font-size:18px; cursor:pointer;">הצג דוח</button>
    </div>
    <script>
        function submitDate() {
            var val = document.getElementById('datePicker').value;
            if(!val) return alert('נא לבחור תאריך');
            var p = val.split('-');
            var formatted = p[2] + '.' + p[1] + '.' + p[0];
            window.location.href = '/api/schedule/Amir?date=' + formatted;
        }
    </script>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date')
    try:
        # ניסיון ראשון: קריאה רגילה
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&sheet={target_date}"
        df = pd.read_csv(url, storage_options={'User-Agent': 'Mozilla/5.0'})
    except Exception:
        try:
            # ניסיון שני: עם רווח לפני (בגלל מה שראינו בטאבים שלך)
            url_space = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&sheet=%20{target_date}"
            df = pd.read_csv(url_space, storage_options={'User-Agent': 'Mozilla/5.0'})
        except Exception as e:
            return f"<div dir='rtl'><h1>שגיאה: לא הצלחתי למצוא את הלשונית {target_date}</h1><p>וודא שהגדרת שיתוף ל'כל מי שקיבל את הקישור'.</p></div>", 401

    df = df.fillna('')
    # חיפוש אמיר אהרון בכל הגיליון
    mask = df.apply(lambda row: row.astype(str).str.contains('אמיר אהרון').any(), axis=1)
    rows = df[mask]
    
    if rows.empty:
        return f"<div dir='rtl'><h1>לא נמצא פרויקט לאמיר ב-{target_date}</h1></div>", 404

    r = rows.iloc[0]
    data = {
        "date": target_date,
        "order": str(r.iloc[0]),   # עמודה A
        "project": str(r.iloc[3]), # עמודה D
        "address": str(r.iloc[4]), # עמודה E
        "inspector": "אמיר אהרון",
        "f_max": 1693.68,
        "sec_e": 1693.68
    }

    if request.args.get('download') == 'excel':
        return generate_excel(data)

    return render_template('index.html', **data)

def generate_excel(data):
    template_path = 'template.xlsx'
    wb = load_workbook(template_path)
    ws = wb.active
    ws['B2'], ws['E2'], ws['B3'], ws['E3'], ws['B4'] = data['date'], data['inspector'], data['project'], data['order'], data['address']
    ws['F15'], ws['F20'] = data['f_max'], data['sec_e']
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Report_{data['date']}.xlsx")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
