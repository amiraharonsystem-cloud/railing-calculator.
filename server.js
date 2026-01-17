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

        // עמוד 1: זיהוי ונתוני שטח
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; 
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassEngStatus;

        // עמוד 2: מפרט טכני וזיגוג (מה שהיה חסר)
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I68').value = d.glassMark;
        ws.getCell('I70').value = d.glassType;
        ws.getCell('I72').value = d.glassManuf;
        ws.getCell('I75').value = d.overlap;
        ws.getCell('I76').value = d.sealing;

        // ערכים לחישוב ועומסי רוח (שורות 80-85)
        ws.getCell('I83').value = d.h_pressure;
        ws.getCell('I84').value = d.qb_val;
        ws.getCell('G82').value = d.we_val;
        ws.getCell('L31').value = d.S_area; 
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;

        // תוצאות בדיקה (סעיף 10.3.4 ו-10.3.5)
        const mapping = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'db': 114, 'e': 116 };
        for (let key in mapping) {
            ws.getCell(`I${mapping[key]}`).value = d[`hor_${key}`];
            ws.getCell(`J${mapping[key]}`).value = d[`res_${key}`];
            ws.getCell(`L${mapping[key]}`).value = d[`stat_${key}`];
        }

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("Error"); }
});
app.listen(10000);
