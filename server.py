import math
import os
import requests
from fpdf import FPDF

# פונקציה לעיבוד טקסט עברי (היפוך סימנים)
def fix_heb(text):
    if not text: return ""
    return str(text)[::-1]

class RailingReport(FPDF):
    def header(self):
        # טעינת פונט עברי מהרשת כדי שיעבוד על כל שרת ללא התקנה
        font_url = "https://github.com/google/fonts/raw/main/ofl/arimo/Arimo%5Bwght%5D.ttf"
        font_path = "HebrewFont.ttf"
        if not os.path.exists(font_path):
            r = requests.get(font_url)
            with open(font_path, 'wb') as f:
                f.write(r.content)
        
        self.add_font("Arimo", "", font_path)
        self.set_font("Arimo", size=16)
        self.cell(0, 10, fix_heb("דוח הנדסי - בדיקת מעקות לפי ת\"י 1142 ו-414"), ln=True, align='C')
        self.ln(10)

def generate_pdf(data):
    pdf = RailingReport()
    pdf.add_page()
    pdf.set_font("Arimo", size=12)
    
    # נתונים מהחישוב
    fw = data.get('fw', 0)
    fser = fw * 0.943
    f_max = max(fw, fser)
    l1 = data.get('l1', 1.0)
    l2 = data.get('l2', 1.0)
    proj_name = data.get('project', 'ללא שם')

    # כתיבת תוכן הדוח
    pdf.cell(0, 10, f"{fix_heb('שם הפרויקט')}: {fix_heb(proj_name)}", ln=True, align='R')
    pdf.cell(0, 10, f"{fix_heb('עומס רוח מחושב')} (Fw): {fw:.2f} N/m", ln=True, align='R')
    pdf.cell(0, 10, f"{fix_heb('עומס שירות')} (Fser): {fser:.2f} N/m", ln=True, align='R')
    pdf.set_text_color(0, 0, 255) # כחול לערך הקובע
    pdf.cell(0, 10, f"{fix_heb('ערך קובע לחישובים')}: {f_max:.2f} N/m", ln=True, align='R')
    pdf.set_text_color(0, 0, 0)
    
    pdf.ln(5)
    pdf.cell(0, 5, "----------------------------------------------------------------", ln=True, align='C')
    pdf.ln(5)

    # סעיפי הבדיקה
    sections = [
        (fix_heb("סעיף א' - עומס על המאחז"), f"{0.375 * l1 * f_max:.2f} N"),
        (fix_heb("סעי
