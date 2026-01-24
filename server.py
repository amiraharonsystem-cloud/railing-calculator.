import os
import io
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# קישור ה-CSV של הגליון שלך מגוגל
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_2H0G6j8iS_F8W6I8-09Gz9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H1z6v9H/pub?output=csv"

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # 1. משיכת נתונים מהגליון
        df = pd.read_csv(SHEET_CSV_URL)
        row = df[(df['בודק'] == 'אמיר אהרון') & (df['תאריך'] == target_date)]
        
        if row.empty:
            return render_template('index.html', error=f"לא נמצא פרויקט לתאריך {target_date}")

        r = row.iloc[0]
        
        # 2. חישובים הנדסיים (זהים לאקסל)
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
            "sec_a": round(0.375 * l1 * f_max, 2),
            "sec_b": round(0.75 * l1 * f_max, 2),
            "sec_c": round(1.2 * f_max, 2),
            "sec_e": round(f_max * l1, 2)
        }

        # 3. אם התבקשה הורדת אקסל
        if request.args.get('download') == 'excel':
            return generate_excel_response(data)

        return render_template('index.html', **data)
    
    except
