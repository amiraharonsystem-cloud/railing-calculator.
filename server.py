import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# הקישור ל-CSV של הגליון שלך
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/api/schedule/Amir')
def get_report():
    try:
        target_date = request.args.get('date', '25/01/2026')
        df = pd.read_csv(SHEET_CSV_URL)
        
        # סינון בודק ותאריך
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return render_template('index.html', error=f"לא נמצא פרויקט לתאריך {target_date}")

        r = row.iloc[0]
        
        # חישובים הנדסיים לפי נוסחאות האקסל שסיפקת
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        f_max = max(fw, fw * 0.943)

        # מיפוי נתונים - שמות המשתנים תואמים למקומם באקסל
        data = {
            "cell_B2": target_date,              # תאריך
            "cell_B3": r.get('שם המזמין', ''),   # שם הפרויקט
            "cell_B4": r.get('כתובת האתר', ''),  # כתובת
            "cell_E3": r.get('מספר הזמנה', ''),  # מספר הזמנה
            "cell_E2": "אמיר אהרון",             # שם הבודק
            "f_max": round(f_max, 2),
            "sec_a": round(0.375 * l1 * f_max, 2),
            "sec_b": round(0.75 * l1 * f_max, 2),
            "sec_c": round(1.2 * f_max, 2),
            "sec_e": round(f_max * l1, 2)         # חישוב סעיף ה'
        }

        return render_template('index.html', **data)
    
    except Exception as e:
        return f"שגיאת מערכת: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
