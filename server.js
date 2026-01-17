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

        // --- עמוד 1: זיהוי ופרטים כלליים ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType;
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassStatus;

        // --- עמוד 2: מפרט חומרים (פרופילים, ברגים, זיגוג) ---
        ws.getCell('C45').value = d.profileType; // סוג הפרופיל
        ws.getCell('F45').value = d.profileFinish; // גמר
        ws.getCell('C47').value = d.anchorType; // סוג עיגון
        ws.getCell('I47').value = d.anchorDesc; // תיאור העיגון
        ws.getCell('C49').value = d.screwType; // ברגים
        ws.getCell('C55').value = d.topRailType; // אזן עליון
        
        // מפרט זכוכית (שורות 66-78)
        ws.getCell('C66').value = d.glassW;
        ws.getCell('F66').value = d.glassH;
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I68').value = d.glassMark;
        ws.getCell('I70').value = d.glassType;
        ws.getCell('I75').value = d.overlap;
        ws.getCell('I76').value = d.sealing;

        // --- עמוד 3: נתונים הנדסיים ותוצאות ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('I32').value = d.We_wind;
        ws.getCell('K32').value = d.S_area;
        ws.getCell('L32').value = d.Ws_total;

        // טבלת תוצאות (107-117)
        const resMap = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'db': 114, 'e': 116 };
        for (let k in resMap) {
            ws.getCell(`I${resMap[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${resMap[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${resMap[k]}`).value = d[`stat_${k}`];
        }

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("Error"); }
});
app.listen(10000);
