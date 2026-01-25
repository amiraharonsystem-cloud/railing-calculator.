import os, io
from flask import Flask, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

app = Flask(__name__)
CORS(app)

def write_to_cell(ws, cell_address, value):
    cell = ws[cell_address]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell_address in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    cell.value = value

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # שתילת הנתונים - תאריך ופרטי פרויקט
        write_to_cell(ws, 'B2', data.get('date', ''))      # תאריך בדיקה
        write_to_cell(ws, 'E2', "אמיר אהרון")             # בודק (קבוע)
        
        write_to_cell(ws, 'B3', data.get('site_name', '')) # שם אתר
        write_to_cell(ws, 'E3', data.get('project', ''))   # קוד פרויקט (בלי IO)
        
        write_to_cell(ws, 'B4', data.get('address', ''))   # כתובת
        write_to_cell(ws, 'E4', data.get('customer', ''))  # שם המזמין
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
