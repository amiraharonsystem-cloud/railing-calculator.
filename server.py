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
    # דף בית עם לוח שנה מעוצב
    return """
    <div dir="rtl" style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; text-align:center; padding:50px; background:#f0f4f8; min-height:100vh;">
        <div style="background:white; padding:30px; border-radius:15px; box-shadow:0 4px 15px rgba(0,0,0,0.1); display:inline-block;">
            <h1 style="color:#1a4e8a; margin-bottom:10px;">מערכת דוחות - אמיר אהרון</h1>
            <p style="font-size:18px; color:#555;">בחר תאריך לביצוע הדוח:</p>
            
            <input type="date" id="calendar" style="padding:12px; font-size:18px; border-radius:8px; border:2px solid #1a4e8a; margin-bottom:20px; width:250px;">
            <br>
            <button onclick="loadByCalendar()" style="padding:12px 30px; font-size:18px; background:#1a4e8a; color:white; border:none; border-radius:8px; cursor:pointer; font-weight:bold; transition:0.3s;">
                הצג נתונים מהגליון
            </button>
        </div>

        <script>
            function loadByCalendar() {
                var rawDate = document.getElementById('calendar').value; // פורמט YYYY-MM-DD
                if(!rawDate) return alert('נא לבחור תאריך');
                
                // המרה לפורמט של הלשוניות שלך (DD.MM.YYYY)
                var parts = rawDate.split('-');
                var formattedDate = parts[2] + '.' + parts[1] + '.' + parts[0];
                
                window.location.href = '/api/schedule/Amir?date=' + encodeURIComponent(formattedDate);
            }
        </script>
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date') # יגיע כ-25.01.2026
    try:
        # קישור ייצוא ללשונית הספציפית
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={target_date}"
        
        df = pd.read_csv(url)
        # חיפוש "אמיר אהרון" בעמודה השנייה (אינדקס 1)
        mask = df.iloc[:, 1].astype(str).str.contains('אמיר אהרון', na=False)
        rows = df[mask]
        
        if rows.empty:
            return f"<div dir='rtl' style='text-align:center; padding:50px;'><h1>לא נמצא פרויקט לאמיר בלשונית {target_date}</h1><a href='/'>חזור לבחירה</a></div>", 404

        r = rows.iloc[0]
        
        # מיפוי נתונים מהלשונית (לפי המבנה שראינו בתמונה)
        data = {
            "date": target_date,
            "order": str(r.iloc[0]),   # עמודה A
            "project": str(r.iloc[3]), # עמודה D
            "address": str(r.iloc[4]), # עמודה E
            "inspector": "אמיר אהרון",
            "f_max": 1693.68, # ערכי ברירת מחדל
            "sec_e": 1693.68
        }

        if request.args.get('download') == 'excel':
            return generate_excel(data)

        return render_template('index.html', **data)
    
    except Exception as e:
        return f"<div dir='rtl' style='padding:50px;'><h1>שגיאה בגישה ללשונית {target_date}</h1><p>וודא ששם הלשונית בגליון תואם בדיוק לתאריך שבחרת ושהגליון מוגדר כ'ציבורי'.</p><a href='/'>חזור לניסיון נוסף</a></div>", 500

def generate_excel(data):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path): return "Template missing", 404
    
    wb = load_workbook(template_path)
    ws = wb.active
    
    # הזרקה לתבנית האקסל המקורית
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
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Report_{data['date']}.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
