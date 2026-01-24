import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/api/schedule/<name>')
def generate_from_excel(name):
    try:
        # קריאת קובץ האקסל שנמצא ב-GitHub שלך
        excel_file = 'template.xlsx'
        df = pd.read_excel(excel_file)

        # נניח שאנחנו מחפשים את השורה של הפרויקט לפי שם הבודק
        # כאן המערכת הופכת ל"זהה לאקסל" - היא שולפת נתונים אמיתיים
        project_data = df[df['בודק'] == name].to_dict(orient='records')
        
        if not project_data:
            # אם לא מצאנו באקסל, נשתמש בנתונים מה-URL כגיבוי
            fw = float(request.args.get('fw', 1693))
            l1 = float(request.args.get('l1', 1.0))
        else:
            # משיכת נתונים ישירות מהאקסל
            fw = project_data[0].get('עומס', 1693)
            l1 = project_data[0].get('מרווח', 1.0)

        # חישוב סעיף ה'
        section_e = fw * l1 # דוגמה לחישוב

        return render_template('index.html', 
                               name=name, 
                               f_max=fw, 
                               section_e=round(section_e, 2))
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500
