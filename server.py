import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# מזהה הגליון שלך מהכתובת שסיפקת
SPREADSHEET_ID = "1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs"

@app.route('/')
def home():
    # דף בית עם תיבת טקסט להזנת תאריך (בדיוק כמו שמות הלשוניות בגליון)
    return """
    <div dir="rtl" style="font-family:Arial; text-align:center; padding:50px; background:#f4f7f6;">
        <h1 style="color:#1a4e8a;">מערכת דוחות הנדסיים - אמיר אהרון</h1>
        <p style="font-size:18px;">הזן תאריך בדיוק כפי שהוא מופיע בלשונית (לדוגמה: 25.01.2026):</p>
        <input type="text" id="dateInput" placeholder="DD.MM.YYYY" style="padding:10px; font-size:18px; border-radius:5px; border:1px solid #ccc;">
        <br><br>
        <button onclick="loadReport()" style="background:#1a4e8a; color:white; padding:10px 30px; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">הצג דוח</button>
        
        <script>
            function loadReport() {
                var date = document.getElementById('dateInput').value;
                if(!date) return alert('נא להזין תאריך');
                window.location.href = '/api/schedule/Amir?date=' + encodeURIComponent(date);
            }
        </script>
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date') # לדוגמה: 25.01.2026
    if not target_date:
        return "נא לספק תאריך", 400

    try:
        # בניית קישור שמייצא לשונית ספציפית לפי השם שלה (target_date)
        # זה מאפשר לנו לעבור בין תאריכים שונים בגליון
        export_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={target_date}"
        
        df = pd.read_csv(export_url)
        df.columns = df.columns.str.strip()
        
        # חיפוש "אמיר אהרון" בעמודה המתאימה (לפי התמונה מדובר בעמודה הראשונה או 'Unnamed: 0')
        # אנחנו נחפש בכל העמודות שורה שמכילה את השם שלו
        mask = df.apply(lambda row: row.astype(str).str.contains('אמיר אהרון').any(), axis=1)
        relevant_rows = df[mask]
        
        if relevant_rows.empty:
            return f"<h1 dir='rtl'>לא נמצא פרויקט לאמיר בתאריך {target_date}</h1><p dir='rtl'>וודא ששם הלשונית תואם בדיוק לתאריך שהזנת.</p><a href='/'>חזור</a>", 404

        r = relevant_rows.iloc[0]
        
        # מיפוי נתונים מהגליון (לפי המבנה ב-image_0887ee.jpg)
        data = {
            "date": target_date,
            "project": str(r.get('שם המזמין', r.iloc[3])), # עמודה D בתמונה
            "address": str(r.get('כתובת האתר', r.iloc[4])), # עמודה E בתמונה
            "order": str(r.get('מספר הזמנה', r.iloc[0])),   # עמודה A בתמונה
            "inspector": "אמיר אהרון",
            "f_max": 1693.68,
            "sec_e": 1693.68
        }

        if request.args.get('download') == 'excel':
            return generate_excel(data)

        return render_template('index.html', **data)
    
    except Exception as e:
        return f"<div dir='rtl'><h1>שגיאת גישה</h1><p>השרת לא מצליח לקרוא את הגליון. וודא שהגדרת 'שיתוף > כל מי שקיבל את הקישור' (Anyone with the link).</p><p>פרטים: {str(e)}</p></div>", 500

def generate_excel(data):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path): return "Template missing", 404
    
    wb = load_workbook(template_path)
    ws = wb.active
    
    # הזרקה למקומות המדויקים באקסל
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
