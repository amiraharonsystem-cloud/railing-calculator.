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

        // --- עמוד 1: זיהוי ומהות ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testKind; // ראשונה / חוזרת
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;

        // --- הצהרות (הוגש / לא הוגש) ---
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassEngStatus;

        // --- ערכים לחישוב (סעיף 3 בקובץ) ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;

        // --- חישוב עומס רוח על לוח מילוי ---
        ws.getCell('L30').value = d.We; // לחץ רוח תכנוני
        ws.getCell('L31').value = d.S_area; // שטח הלוח
        ws.getCell('L32').value = d.Ws_total; // Ws = We * S

        // --- תוצאות בדיקה (10.3.4 ו-10.3.5) ---
        ws.getCell('I107').value = d.hor_a; ws.getCell('J107').value = d.res_a; ws.getCell('L107').value = d.stat_a;
        ws.getCell('I108').value = d.hor_b; ws.getCell('J108').value = d.res_b; ws.getCell('L108').value = d.stat_b;
        ws.getCell('I110').value = d.hor_c; ws.getCell('J110').value = d.res_c; ws.getCell('L110').value = d.stat_c;

        // --- חתימות וסיכום ---
        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.testerName;
        ws.getCell('L39').value = d.approverName;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Report_Complete.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("שגיאה בהפקה"); }
});

app.listen(10000, '0.0.0.0');
