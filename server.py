import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# הקישור לנתוני גוגל שיטס שלך
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # קריאת נתונים מהגליון
        df = pd.read_csv(SHEET_CSV_URL)
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return "לא נמצאו נתונים עבור אמיר אהרון בתאריך זה", 404

        r = row.iloc[0]
        
        # חישובים (כדי להציג ב-HTML לפני ההורדה)
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        f_max = max(fw, fw * 0.943)
        section_e = f_max * l1

        data = {
            "date": target_date,
            "project": r.get('שם המזמין', ''),
            "address": r.get('כתובת האתר', ''),
            "order": r.get('מספר הזמנה', ''),
            "inspector": "אמיר אהרון",
            "f_max": round(f_max, 2),
            "sec_e": round(section_e, 2)
        }

        # אם המשתמש לוחץ על הורדה
        if request.args.get('download') == 'excel':
            return generate_excel_from_template(data)

        return render_template('index.html', **data)
    
    except Exception as e:
        return f"שגיאה בעיבוד הנתונים: {str(e)}", 500

def generate_excel_from_template(data):
    # טעינת הקובץ המדויק מה-Git שלך
    template_path = 'template.xlsx'
    if not os.path.exists(template_path):
        return "קובץ template.xlsx לא נמצא בשרת", 404
        
    wb = load_workbook(template_path)
    ws = wb.active

    # הזרקת הנתונים למקומות המקוריים
    # (כאן המערכת משתמשת במיפוי התאים של האקסל שסיפקת)
    ws['B2'] = data['date']
    ws['E2'] = data['inspector']
    ws['B3'] = data['project']
    ws['E3'] = data['order']
    ws['B4'] = data['address']
    
    # הזרקת חישובים לתאים הרלוונטיים
    # (התאים הללו מותאמים למבנה האקסל ההנדסי)
    ws['F15'] = data['f_max']
    ws['F20'] = data['sec_e']

    # שמירה לזיכרון ושליחה למשתמש
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"Report_Amir_{data['date'].replace('/','-')}.xlsx"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
