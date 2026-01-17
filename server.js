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

        // --- פרטי פרויקט ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.clientRep;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLoc;
        ws.getCell('C13').value = d.planningStatus; 
        ws.getCell('C14').value = d.engStatus;

        // --- פרטי זכוכית ולוחות ---
        ws.getCell('B66').value = d.glassLength;
        ws.getCell('E66').value = d.glassHeight;
        ws.getCell('H67').value = d.glassThick;
        ws.getCell('H68').value = d.glassMarking;
        ws.getCell('H70').value = d.glassType;
        ws.getCell('H72').value = d.glassManufacturer;
        ws.getCell('H75').value = d.overlap;
        ws.getCell('H76').value = d.sealingMaterial;

        // --- עומסי רוח ---
        ws.getCell('C83').value = d.heightH;
        ws.getCell('H82').value = d.qbValue;
        ws.getCell('G82').value = d.weValue;
        ws.getCell('G83').value = d.fwValue;
        ws.getCell('G84').value = d.cezeValue;
        ws.getCell('G85').value = d.topRailLoad;

        // --- טבלת בדיקות 10.3.4 (מדידות ובחירה) ---
        const rows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112 };
        for (let key in rows) {
            let r = rows[key];
            ws.getCell(`I${r}`).value = d[`horiz_${key}`]; // תזוזה אופקית
            ws.getCell(`J${r}`).value = d[`resid_${key}`]; // תזוזה שיורית
            ws.getCell(`L${r}`).value = d[`res_${key}`];   // עמד/לא עמד
        }

        // --- סיכום וחתימה ---
        ws.getCell('C40').value = d.notes;
        ws.getCell('C42').value = d.testerName;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Test_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error"); }
});

app.listen(10000, '0.0.0.0');
