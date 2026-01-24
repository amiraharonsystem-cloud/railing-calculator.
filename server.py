import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# --- כאן תדביק את הקישור שקיבלת מגוגל שיטס (חובה שיהיה בסיומת output=csv) ---
SHEET_CSV_URL = "YOUR_CSV_LINK_HERE"

@app.route('/')
def home():
    return """
    <div dir="rtl" style="font-family:Arial; text-align:center; padding:50px;">
        <h1>המערכת פעילה! 🚀</h1>
        <p>כדי לראות את הדוח של אמיר אהרון למחר, לחץ כאן:</p>
        <a href="/api/schedule/Amir?date=25/01/2026" style="font-size:20px;">צפה בדוח ל-25/01/2026</a>
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # 1. משיכת הנתונים מגוגל שיטס
        df = pd.read_csv(SHEET_CSV_URL)
        
        # ניקוי רווחים משמות העמודות (למניעת שגיאות)
        df.columns = df.columns.str.strip()
        
        # סינון לפי בודק ותאריך
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return f"<h1 dir='rtl'>לא נמצאו נתונים עבור אמיר אהרון בתאריך {target_date}</h1>", 404

        r = row.iloc[0]
        
        # 2. הכנת הנתונים (שימוש בערכי ברירת מחדל אם התא ריק)
        data = {
            "date": target_date,
            "project": str(r.get('שם המזמין', 'ללא שם')),
            "address": str(r.get('כתובת האתר', 'ללא כתובת')),
            "order": str(r.get('מספר הזמנה', '0')),
            "inspector": "אמיר אהרון",
            "fw": 1693.68, # ערך קבוע מהאקסל המקורי
            "l1": 1.0,     # ערך קבוע מהאקסל המקורי
        }
        
        # חישובים הנדסיים זהים לאקסל
        f_max = max(data['fw'], data['fw'] * 0.943)
        data['f_max'] = round(f_max, 2)
        data['sec_e'] = round(f_max * data['l1'], 2)

        # 3. אם המשתמש לוחץ על הורדת אקסל
        if request.args.get('download') == 'excel':
            return generate_excel_response(data)

        # 4. תצוגה באתר
        return render_template('index.html', **data)
    
    except Exception as e:
        return jsonify({"error": "שגיאה בגישה לנתונים. וודא שהקישור לגוגל שיטס תקין.", "details": str(e)}), 500

def generate_excel_response(data):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path):
        return "שגיאה: קובץ template.xlsx לא נמצא ב-GitHub שלך", 404
        
    wb = load_workbook(template_path)
    ws = wb.active

    # הזרקה לתאים בדיוק לפי המבנה המקורי של האקסל שלך
    ws['B2'] = data['date']      # תאריך
    ws['E2'] = data['inspector'] # בודק
    ws['B3'] = data['project']   # מזמין/פרויקט
    ws['E3'] = data['order']     # מס' הזמנה
    ws['B4'] = data['address']   # כתובת האתר
    
    # הזרקת חישובים הנדסיים
    ws['F15'] = data['f_max']    # עומס תכנוני
    ws['F20'] = data['sec_e']    # סעיף ה'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"Report_Amir_{data['date'].replace('/', '_')}.xlsx"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
