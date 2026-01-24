import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# הקישור הישיר ל-CSV של הגליון שלך (לוודא שביצעת "פרסם באינטרנט" כ-CSV)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/')
def home():
    return "Server is Live. Use /api/schedule/Amir"

@app.route('/api/schedule/Amir')
def get_report():
    try:
        # קבלת התאריך מהדפדפן (ברירת מחדל למחר 25/01/2026)
        target_date = request.args.get('date', '25/01/2026')
        
        # קריאת הגליון
        df = pd.read_csv(SHEET_CSV_URL)
        
        # סינון לפי אמיר אהרון והתאריך הנבחר
        # הנחה: עמודות בגליון נקראות 'בודק' ו-'תאריך'
        amir_row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if amir_row.empty:
            return render_template('index.html', error=f"לא נמצא פרויקט לאמיר בתאריך {target_date}")

        # שליפת נתונים מהשורה
        project_info = {
            "name": "אמיר אהרון",
            "date": target_date,
            "project": amir_row.iloc[0].get('שם המזמין', 'לא צוין'),
            "address": amir_row.iloc[0].get('כתובת האתר', 'לא צוין'),
            "order": amir_row.iloc[0].get('מספר הזמנה', '0'),
            "fw": float(request.args.get('fw', 1693.68)), # מהירות רוח/עומס
            "l1": float(request.args.get('l1', 1.0)),
            "l2": float(request.args.get('l2', 1.0))
        }
        
        # חישובים הנדסיים (סעיף ה')
        f_max = max(project_info['fw'], project_info['fw'] * 0.943)
        project_info['section_e'] = round(f_max * (project_info['l1']/2 + project_info['l2']/2), 2)
        project_info['f_max'] = round(f_max, 2)

        return render_template('index.html', **project_info)
    
    except Exception as e:
        return f"Error: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
