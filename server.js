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

        // --- מילוי לפי כתובת תא מדויקת ---
        // פרטי זיהוי
        ws.getCell('H2').value = d.H2;   // מספר דו"ח
        ws.getCell('L3').value = d.L3;   // תאריך
        ws.getCell('C4').value = d.C4;   // שם האתר
        ws.getCell('L4').value = d.L4;   // קוד פרויקט
        ws.getCell('C7').value = d.C7;   // שם המזמין
        ws.getCell('I7').value = d.I7;   // מען המזמין
        ws.getCell('L7').value = d.L7;   // נציג המזמין
        ws.getCell('C8').value = d.C8;   // כתובת האתר
        ws.getCell('D9').value = d.D9;   // מהות הבדיקה

        // מפרט טכני (השורות שביקשת)
        ws.getCell('C11').value = d.C11; // תיאור המבנה
        ws.getCell('C12').value = d.C12; // תיאור הפריט
        ws.getCell('C13').value = d.C13; // מיקום הבדיקה
        ws.getCell('I14').value = d.I14; // תכנון מעקה
        ws.getCell('I15').value = d.I15; // חישוב שלד
        ws.getCell('I16').value = d.I16; // חישוב מליא

        // נתוני חומרים וברגים (שורות 45-78)
        ws.getCell('C45').value = d.C45; // סוג פרופיל
        ws.getCell('F45').value = d.F45; // גמר
        ws.getCell('C47').value = d.C47; // סוג עיגון
        ws.getCell('I47').value = d.I47; // תיאור עיגון
        ws.getCell('C49').value = d.C49; // ברגים
        ws.getCell('C55').value = d.C55; // מסעד יד
        ws.getCell('I67').value = d.I67; // עובי זכוכית
        ws.getCell('I70').value = d.I70; // סוג זכוכית

        // נתוני חישוב ועומסים
        ws.getCell('L19').value = d.L19; // Fser
        ws.getCell('L20').value = d.L20; // P מחושב
        ws.getCell('L22').value = d.L22; // L1
        ws.getCell('L24').value = d.L24; // L2
        ws.getCell('I32').value = d.I32; // We לחץ רוח
        ws.getCell('L32').value = d.L32; // Ws עומס כולל

        // תוצאות
        ws.getCell('I107').value = d.I107; // תזוזה א'
        ws.getCell('J107').value = d.J107; // שיורית א'
        ws.getCell('L107').value = d.L107; // מסקנה א'

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send(e.message); }
});

app.listen(process.env.PORT || 3000);
