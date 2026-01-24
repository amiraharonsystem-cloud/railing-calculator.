import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# קישור ה-CSV של הגליון שלך (Published to web)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # 1. משיכת הנתונים מהגליון (מערכות בניין 2026)
        df = pd.read_csv(SHEET_CSV_URL)
        # סינון לפי בודק ותאריך
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return render_template('index.html', error=f"לא נמצאו נתונים עבור אמיר בתאריך {target_date}")

        r = row.iloc[0]
        
        # 2. איסוף נתונים לחישוב (מתוך האתר או הגליון)
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        f_max = max(fw, fw * 0.943)

        data = {
            "date": target_date,
            "project": r.get('שם המזמין', 'חסר נתון'),
            "address": r.get('כתובת האתר', 'חסר נתון'),
            "order": r.get('מספר הזמנה', '0'),
            "inspector": "אמיר אהרון",
            "f_max": round(f_max, 2),
            "sec_a": round(0.375 * l1 * f_max, 2),
            "sec_b": round(0.75 * l1 * f_max, 2),
            "sec_c": round(1.2 * f_max, 2),
            "sec_e": round(f_max * l1, 2)
        }

        # 3. אם המשתמש לוחץ על הורדת אקסל
        if request.args.get('download') == 'excel':
            return generate_excel_response(data)

        # 4. תצוגה דינמית ב-HTML
        return render_template('index.html', **data)
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def generate_excel_response(data):
    # טעינת קובץ המקור (התבנית שלך)
    template_path = 'template.xlsx'
    if not os.path.exists(template_path):
        return "קובץ template.xlsx לא נמצא בשרת GitHub שלך", 404
        
    wb = load_workbook(template_path)
    ws = wb.active

    # הזרקה לתאים - התאמה מדויקת למבנה האקסל המקורי
    ws['B2'] = data['date']      # תאריך
    ws['E2'] = data['inspector'] # בודק
    ws['B3'] = data['project']   # מזמין/פרויקט
    ws['E3'] = data['order']     # מס' הזמנה
    ws['B4'] = data['address']   # כתובת האתר
    
    # הזרקת חישובים הנדסיים לתאים הרלוונטיים
    ws['F15'] = data['f_max']    # עומס תכנוני
    ws['F17'] = data['sec_a']    # סעיף א'
    ws['F18'] = data['sec_b']    # סעיף ב'
    ws['F19'] = data['sec_c']    # סעיף ג'
    ws['F20'] = data['sec_e']    # סעיף ה'

    # שמירה לזיכרון ושליחה למשתמש
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Report_Amir_{data['date'].replace('/', '-')}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
