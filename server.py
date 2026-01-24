import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# הקישור הישיר לגליון שלך (מותאם לפורמט CSV)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs/gviz/tq?tqx=out:csv&gid=1101827105"

@app.route('/')
def home():
    return """
    <div dir="rtl" style="font-family:Arial; text-align:center; padding:50px; background-color:#f4f7f6;">
        <h1 style="color:#1a4e8a;">מערכת דוחות הנדסיים - אמיר אהרון</h1>
        <p style="font-size:18px;">המערכת מחוברת לגליון הנתונים ומכילה את תבנית האקסל המקורית.</p>
        <hr style="width:50%; margin:20px auto;">
        <p>כדי לצפות בדוח למחר:</p>
        <a href="/api/schedule/Amir?date=25/01/2026" 
           style="display:inline-block; background:#1a4e8a; color:white; padding:15px 30px; text-decoration:none; border-radius:5px; font-weight:bold;">
           צפה בדוח ל-25/01/2026
        </a>
    </div>
    """

@app.route('/api/schedule/Amir')
def get_report():
    target_date = request.args.get('date', '25/01/2026')
    try:
        # 1. משיכת הנתונים מהגליון
        df = pd.read_csv(SHEET_CSV_URL)
        
        # ניקוי
