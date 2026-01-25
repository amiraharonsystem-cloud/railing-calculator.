import os
import io
from flask import Flask, request, send_file
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
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
        template_path = 'template.xlsx'
        
        if os.path.exists(template_path):
            wb = load_workbook(template_path)
        else:
            wb = Workbook()
            
        ws = wb.active
        
        # שימוש בפונקציה החדשה שמונעת את שגיאת ה-MergedCell
        write_to_cell(ws, 'B2', data.get('date', ''))
        write_to_cell(ws, 'F2', "אמיר אהרון")
        write_to_cell(ws, 'B3', data.get('project', ''))
        write_to_cell(ws, 'E3', data.get('order', ''))
        write_to_cell(ws, 'B4', data.get('address', ''))

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name="Report.xlsx")
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
