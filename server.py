from flask import Flask, request, send_file
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import io
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.json
    cells = data.get('cells', {})
    # מקבל את שם הקובץ הדינמי שנשלח מה-Frontend
    file_name = data.get('fileName', 'Report_1142.xlsx')
    
    # נתיב לקובץ הטמפלייט המקורי שלך
    template_path = 'template.xlsx' 
    
    if not os.path.exists(template_path):
        return "Template file not found", 404

    # טעינת האקסל המקורי
    wb = load_workbook(template_path)
    ws = wb.active # או wb['שם הגיליון']

    # כתיבת הנתונים לתאים לפי הכתובות ששלחנו מה-HTML
    for addr, value in cells.items():
        try:
            # המרה למספר אם ניתן, כדי שהאקסל יוכל לבצע חישובים
            if value and value.replace('.', '', 1).isdigit():
                ws[addr] = float(value)
            else:
                ws[addr] = value
        except Exception as e:
            print(f"Error writing to cell {addr}: {e}")

    # שמירה לזיכרון ושליחה למשתמש
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    return send_file(
        excel_file,
        as_attachment=True,
        download_name=file_name, # משתמש בשם הקובץ הדינמי
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True, port=5000)
