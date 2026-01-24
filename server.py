import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# מזהה הגיליון שלך
SPREADSHEET_ID = "1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs"

@app.route('/')
def home():
    return """
    <div dir="rtl" style="font-family:Arial, sans-serif; text-align:center; padding:50px; background:#f4f7f8; min-height:100vh;">
        <div style="background:white; padding:30px; border-radius:15px; box-shadow:0 10px 25px rgba(0,0,0,0.1); display:inline-block;">
            <h1 style="color:#1a4e8a;">מערכת דוחות הנדסיים</h1>
            <p style="font-size:18px;">בחר תאריך מהלוח להפקת דוח:</p>
            <input type="date" id="datePicker" style="padding:15px; font-size:18px; border-radius:8px; border:2px solid #1a4e8a; width:250px;">
            <br><br>
            <button onclick="submitDate()" style="background:#1a4e8a; color:white; padding:15px 40px; border:none; border-radius:8px; cursor:pointer; font-weight:bold; font-size:18px;">
                חפש בגיליון והצג דוח
            </button>
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
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date') 
    try:
        # פתרון חסין: אנחנו מנסים למשוך את הלשונית לפי השם שנבחר
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={target_date}"
        df = pd.read_csv(url)
        
        # ניקוי רווחים מהנתונים
        df = df.fillna('')
        
        # חיפוש "אמיר אהרון" בכל הגיליון (למקרה שהעמודה זזה)
        mask = df.apply(lambda row: row.astype(str).str.contains('אמיר אהרון').any(), axis=1)
        rows = df[mask]
        
        if rows.empty:
             return f"<div dir='rtl' style='text-align:center; padding:50px;'><h1>לא נמצא פרויקט לאמיר בתאריך {target_date}</h1><p>וודא שהלשונית קיימת בגיליון בשם זה.</p><a href='/'>חזור ללוח השנה</a></div>", 404

        r = rows.iloc[0]
        
        # שליפת נתונים לפי מיקום (מבוסס על הצילום שלך)
        data = {
            "date": target_date,
            "order": str(r.iloc[0]) if len(r) > 0 else "0",   # עמודה A
            "project": str(r.iloc[3]) if len(r) > 3 else "חסר", # עמודה D
            "address": str(r.iloc[4]) if len(r) > 4 else "חסר", # עמודה E
            "inspector": "אמיר אהרון",
            "f_max": 1693.68,
            "sec_e": 1693.68
        }

        if request.args.get('download') == 'excel':
            return generate_excel(data)

        # הוספת כפתור "חזור" בדף התוצאה
        return f'''
        {render_template('index.html', **data)}
        <div style="text-align:center; margin-top:20px;">
            <a href="/" style="text-decoration:none; color:#1a4e8a; font-weight:bold;">← חזור לבחירת תאריך אחר</a>
        </div>
        '''
    
    except Exception as e:
        return f"<div dir='rtl' style='padding:50px;'><h1>שגיאה בגישה ללשונית {target_date}</h1><p>וודא שהגדרת שיתוף ל'כל מי שקיבל את הקישור' וששם הלשונית מדויק.</p><a href='/'>נסה שוב</a></div>", 500

def generate_excel(data):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path): return "Template missing", 404
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
