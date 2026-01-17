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

        // --- פרטי זיהוי (3-14) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.clientRep; // שם המציג / נציג המזמין
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structure;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLoc;
        ws.getCell('C13').value = d.planning;
        ws.getCell('C14').value = d.engineering;

        // --- נתוני מדידה (22-31) ---
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L30').value = d.ws;
        ws.getCell('L31').value = d.half_ws;

        // --- לוחות מילוי וזכוכית (66-77) ---
        ws.getCell('B66').value = d.glassLength; // אורך דיגוז
        ws.getCell('E66').value = d.glassHeight; // גובה דיגוז
        ws.getCell('H67').value = d.glassThick;  // עובי במ"מ
        ws.getCell('H68').value = d.manufacturer; // סימון יצרן
        ws.getCell('H70').value = d.glassType;   // סוג שמשה
        ws.getCell('H72').value = d.glassBrand;  // שם יצרן
        ws.getCell('H75').value = d.overlap;     // חפיפה
        ws.getCell('H76').value = d.sealing;     // חומר אטם

        // --- חישוב עומסי רוח (80-85) ---
        ws.getCell('C83').value = d.h;     // גובה הפעלת לחץ
        ws.getCell('H82').value = d.qb;    // qb
        ws.getCell('G82').value = d.We;    // We
        ws.getCell('G83').value = d.Fw;    // Fw
        ws.getCell('G84').value = d.CeZe;  // Ce(Ze)
        ws.getCell('G85').value = d.topRailWind; // עומס רוח אזן עליון

        // --- טבלת בדיקות ותזוזות (107-116) ---
        const testRows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let key in testRows) {
            let r = testRows[key];
            ws.getCell(`I${r}`).value = d[`I${r}`]; // תזוזה אופקית
            ws.getCell(`J${r}`).value = d[`J${r}`]; // תזוזה שיורית
            ws.getCell(`L${r}`).value = d[`L${r}`]; // עמד/לא עמד
        }

        // --- סיכום (40-42) ---
        ws.getCell('C40').value = d.comments;
        ws.getCell('C42').value = d.inspector;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Full_Railing_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error"); }
});

app.listen(10000, '0.0.0.0');
