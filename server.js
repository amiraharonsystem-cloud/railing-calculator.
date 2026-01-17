const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // --- עמוד 1: זיהוי וכללי ---
        ws.getCell('H2').value = d.H2;   // מס' דוח
        ws.getCell('L3').value = d.L3;   // תאריך
        ws.getCell('C4').value = d.C4;   // שם אתר
        ws.getCell('C7').value = d.C7;   // מזמין
        ws.getCell('I14').value = d.I14; // תכנון הוגש (V/X)

        // --- עמוד 2: גאומטריה ופרטים (שורות 55-78) ---
        ws.getCell('E56').value = d.E56; // גובה מעקה
        ws.getCell('E60').value = d.E60; // מרווחים בין רכיבים
        ws.getCell('E66').value = d.E66; // אורך זיגוג
        ws.getCell('E67').value = d.E67; // גובה זיגוג
        ws.getCell('G67').value = d.G67; // עובי זכוכית
        ws.getCell('G70').value = d.G70; // סוג זכוכית
        ws.getCell('G75').value = d.G75; // עובי שכבה פנימית

        // --- עמוד 2: חישובי רוח (שורות 79-85) ---
        ws.getCell('F83').value = d.F83; // h - גובה מעל פני הקרקע
        ws.getCell('F84').value = d.F84; // qb - לחץ רוח בסיסי
        ws.getCell('H82').value = d.H82; // We - לחץ רוח לתכנון
        ws.getCell('H84').value = d.H84; // Fw - כוח רוח כולל

        // --- עמוד 3: תוצאות עומסים (שורות 100-138) ---
        // סעיף 10.3.4 א'
        ws.getCell('I107').value = d.I107; // תזוזה אופקית
        ws.getCell('J107').value = d.J107; // תזוזה שיורית
        ws.getCell('L107').value = d.L107; // מסקנה (מתאים/לא)
        
        // סעיף 10.3.5 ג' (מליא)
        ws.getCell('H127').value = d.H127; // S - שטח הלוח
        ws.getCell('I132').value = d.I132; // תוצאת בדיקה מליא
        ws.getCell('L133').value = d.L133; // מסקנה סופית

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send(e.message); }
});

app.listen(process.env.PORT || 3000);
