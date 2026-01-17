const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: פרטי זיהוי (שורות 3-14) ---
        ws.getCell('L3').value = d.L3;   // תאריך
        ws.getCell('C4').value = d.C4;   // שם אתר
        ws.getCell('L4').value = d.L4;   // קוד פרויקט
        ws.getCell('C7').value = d.C7;   // שם מזמין
        ws.getCell('L7').value = d.L7;   // נציג
        ws.getCell('C8').value = d.C8;   // כתובת
        ws.getCell('C10').value = d.C10; // תיאור מבנה
        ws.getCell('C11').value = d.C11; // תיאור פריט
        ws.getCell('C12').value = d.C12; // מיקום בדיקה
        ws.getCell('C13').value = d.C13; // תכנון (הוגש/לא הוגש)
        ws.getCell('C14').value = d.C14; // חישוב (הוגש/לא הוגש)

        // --- חלק 2: גיאומטריה ומדידות (שורות 22-31) ---
        ws.getCell('L22').value = d.L22; // L1
        ws.getCell('L24').value = d.L24; // L2
        ws.getCell('L26').value = d.L26; // L_fill
        ws.getCell('L30').value = d.L30; // ws
        ws.getCell('L31').value = d.L31; // 0.5ws

        // --- חלק 3: תוצאות בדיקה (שורות 35-138) ---
        ws.getCell('G35').value = d.G35; // עמד/לא עמד
        ws.getCell('G36').value = d.G36; // עמד/לא עמד
        ws.getCell('G37').value = d.G37; // עמד/לא עמד
        
        ws.getCell('C40').value = d.C40; // הערות
        ws.getCell('C42').value = d.C42; // שם בודק

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Report_Full.xlsx');
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        res.status(500).send("שגיאה: וודא שקובץ template.xlsx נמצא ב-GitHub");
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0');
