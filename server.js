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

        // זיהוי ואתר
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.clientRep;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structure;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLoc;
        ws.getCell('C13').value = d.planning; 
        ws.getCell('C14').value = d.engineering;

        // נתוני מדידה וגיאומטריה
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L30').value = d.ws;
        ws.getCell('L31').value = d.half_ws;

        // טבלת בדיקות 10.3.4 (כולל תזוזות אופקיות ושיוריות)
        const rows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let key in rows) {
            let r = rows[key];
            ws.getCell(`I${r}`).value = d[`I${r}`]; // תזוזה אופקית
            ws.getCell(`J${r}`).value = d[`J${r}`]; // תזוזה שיורית
            ws.getCell(`L${r}`).value = d[`L${r}`]; // תוצאה: עמד/לא עמד
        }

        // סיכום
        ws.getCell('C40').value = d.comments;
        ws.getCell('C42').value = d.inspector;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error"); }
});

app.listen(10000, '0.0.0.0');
