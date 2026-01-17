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

        // --- זיהוי ואתר (3-14) ---
        ws.getCell('L3').value = d.L3; ws.getCell('C4').value = d.C4;
        ws.getCell('L4').value = d.L4; ws.getCell('C7').value = d.C7;
        ws.getCell('L7').value = d.L7; ws.getCell('C8').value = d.C8;
        ws.getCell('C10').value = d.C10; ws.getCell('C11').value = d.C11;
        ws.getCell('C12').value = d.C12; ws.getCell('C13').value = d.C13;
        ws.getCell('C14').value = d.C14;

        // --- מדידות (22-31) ---
        ws.getCell('L22').value = d.L22; ws.getCell('L24').value = d.L24;
        ws.getCell('L26').value = d.L26; ws.getCell('L30').value = d.L30;
        ws.getCell('L31').value = d.L31;

        // --- טבלת בדיקות מפורטת (שורות 103-122 כפי שצילמת) ---
        const rows = ['107', '108', '110', '112', '114', '116'];
        rows.forEach(r => {
            ws.getCell(`I${r}`).value = d[`I${r}`]; // תזוזה אופקית
            ws.getCell(`J${r}`).value = d[`J${r}`]; // תזוזה שיורית
            ws.getCell(`L${r}`).value = d[`L${r}`]; // תוצאה: עמד/לא עמד
        });

        // --- סיכום (40-42) ---
        ws.getCell('C40').value = d.C40; ws.getCell('C42').value = d.C42;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Report_Full.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send(e.message); }
});

app.listen(10000, '0.0.0.0');
