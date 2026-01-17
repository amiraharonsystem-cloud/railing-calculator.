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

        // --- עמוד 1: זיהוי ופרטים כלליים ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType;
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;
        ws.getCell('I14').value = d.planOk;
        ws.getCell('I15').value = d.engOk;
        ws.getCell('I16').value = d.glassOk;

        // --- עמוד 2: מפרט טכני מפורט (חומרים, עיגון, זכוכית) ---
        ws.getCell('C45').value = d.profileType;   // סוג פרופיל
        ws.getCell('F45').value = d.profileFinish; // גמר
        ws.getCell('C47').value = d.anchorType;    // סוג עיגון
        ws.getCell('I47').value = d.anchorDesc;    // תיאור עיגון
        ws.getCell('C49').value = d.screwType;     // ברגים
        ws.getCell('C55').value = d.topRail;       // מסעד יד
        ws.getCell('C66').value = d.glassW;        // רוחב זכוכית
        ws.getCell('F66').value = d.glassH;        // גובה זכוכית
        ws.getCell('I67').value = d.glassThick;    // עובי זכוכית
        ws.getCell('I70').value = d.glassType;     // סוג זכוכית
        ws.getCell('I75').value = d.overlap;       // חפיפה
        ws.getCell('I76').value = d.sealing;       // אטימה

        // --- עמוד 2: נתוני רוח הנדסיים ---
        ws.getCell('I83').value = d.h_ground;
        ws.getCell('I84').value = d.qb;
        ws.getCell('G82').value = d.ce_z;
        ws.getCell('G83').value = d.cp_net;
        ws.getCell('G84').value = d.we_wind;

        // --- עמוד 3: חישובי עומסים ותוצאות ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('I32').value = d.we_final;
        ws.getCell('K32').value = d.S_area;
        ws.getCell('L32').value = d.ws_total;

        // טבלת תוצאות סופית (שורות 107-116)
        const rows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let k in rows) {
            ws.getCell(`I${rows[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${rows[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${rows[k]}`).value = d[`stat_${k}`];
        }

        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.tester;
        ws.getCell('L39').value = d.approver;

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send(e.message); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
