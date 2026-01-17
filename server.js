const express = require('express');
const ExcelJS = require('exceljs');
const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('template.xlsx');
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: פרטי זיהוי (שורות 1-10) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; // "ראשונה" או "חוזרת"

        // --- חלק 2: הצהרות ותכנון (שורות 14-16) ---
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassStatus;

        // --- תיאור המערכת (שורות 11-13) ---
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;

        // --- חלק 3: נתוני בדיקה וערכים לחישוב (סוף ה-CSV) ---
        ws.getCell('L19').value = d.Fser;     // עומס שירות
        ws.getCell('L20').value = d.P_calc;   // עומס מחושב P
        ws.getCell('L22').value = d.L1;       // מפתח L1
        ws.getCell('L24').value = d.L2;       // מפתח L2
        ws.getCell('L26').value = d.L_fill;   // מפתח L (לוח מליא)

        // --- חישוב עומס רוח (סעיף 10.3.5 ג') ---
        ws.getCell('I32').value = d.we_val;   // we
        ws.getCell('K32').value = d.S_area;   // S - שטח הלוח
        ws.getCell('L32').value = d.ws_total; // ws=we*S

        // --- טבלת תוצאות (107, 108, 110 וכו') ---
        ws.getCell('I107').value = d.hor_a; ws.getCell('J107').value = d.res_a; ws.getCell('L107').value = d.stat_a;
        ws.getCell('I108').value = d.hor_b; ws.getCell('J108').value = d.res_b; ws.getCell('L108').value = d.stat_b;
        ws.getCell('I110').value = d.hor_c; ws.getCell('J110').value = d.res_c; ws.getCell('L110').value = d.stat_c;

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("Error"); }
});
app.listen(10000);
