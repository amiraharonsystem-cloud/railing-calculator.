import os
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/api/schedule/<name>', methods=['GET', 'POST'])
def generate_dynamic_report(name):
    try:
        # שליפת נתוני החישוב מהכתובת (URL parameters)
        # אם הנתונים לא קיימים, נשתמש בערכי ברירת מחדל
        fw = float(request.args.get('fw', 1693.68))
        l1 = float(request.args.get('l1', 1.0))
        l2 = float(request.args.get('l2', 1.0))
        
        # ביצוע החישוב ההנדסי (סעיף ה')
        f_max = max(fw, fw * 0.943)
        section_e = f_max * (l1/2 + l2/2)

        # פקודת הקסם: מחברת את הנתונים לקובץ ה-HTML שיצרת בתיקיית templates
        return render_template('index.html', 
                               name=name, 
                               f_max=round(f_max, 2), 
                               section_e=round(section_e, 2))
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
