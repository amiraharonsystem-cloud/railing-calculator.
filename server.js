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

        // --- חלק 1: כותרות וזיהוי (שורות 2-10) ---
        ws.getCell('H2').value = d.reportNum;
        ws.getCell('L2').value = d.replacedReport;
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType;

        // --- חלק 2: מפרט טכני של המעקה (שורות 11-78) ---
        ws.getCell('C11').value = d.structDesc;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassStatus;
        
        ws.getCell('C45').value = d.profileType;
        ws.getCell('F45').value = d.profileFinish;
        ws.getCell('C47').value = d.anchorType;
        ws.getCell('I47').value = d.anchorDesc;
        ws.getCell('C49').value = d.screwType;
        ws.getCell('C55').value = d.topRail;
        ws.getCell('C66').value = d.glassW;
        ws.getCell('F66').value = d.glassH;
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I70').value = d.glassType;

        // --- חלק 3: נתוני רוח ועומסים (שורות 19-85) ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('I83').value = d.h_ground;
        ws.getCell('I84').value = d.qb;
        ws.getCell('G84').value = d.we_design;
        ws.getCell('L32').value = d.ws_total;

        // --- חלק 4: טבלת תוצאות מלאה (107-117) ---
        const rows = { '107':'a', '108':'b', '110':'c', '112':'d', '114':'db', '116':'e' };
        for (let r in rows) {
            ws.getCell(`I${r}`).value = d[`hor_${rows[r]}`];
            ws.getCell(`J${r}`).value = d[`res_${rows[r]}`];
            ws.getCell(`L${r}`).value = d[`stat_${rows[r]}`];
        }

        // --- חתימות (37-40) ---
        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.tester;
        ws.getCell('L39').value = d.approver;

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send(e.message); }
});

app.listen(process.env.PORT || 3000);
