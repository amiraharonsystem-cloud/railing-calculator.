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

        // --- פרטי זיהוי וסביבה (3-16) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structureType;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLocation;
        ws.getCell('C13').value = d.planningStatus;
        ws.getCell('C14').value = d.engineeringStatus;
        ws.getCell('L15').value = d.temp; // טמפרטורה
        ws.getCell('L16').value = d.humidity; // לחות

        // --- התאמה לסעיפי תקן (28-31) ---
        ws.getCell('L28').value = d.match1;
        ws.getCell('L29').value = d.match2;
        ws.getCell('L30').value = d.match3;
        ws.getCell('L31').value = d.match4;

        // --- גאומטריה ועומסים (הזרקה כפולה 22-31 ו-118-122) ---
        ws.getCell('L22').value = d.L1; ws.getCell('L118').value = d.L1;
        ws.getCell('L24').value = d.L2; ws.getCell('L120').value = d.L2;
        ws.getCell('L26').value = d.L_fill; ws.getCell('L122').value = d.L_fill;
        ws.getCell('L30').value = d.ws_load;
        ws.getCell('L31').value = d.half_ws;

        // --- מפרט זכוכית (66-77) ---
        ws.getCell('C66').value = d.glassW;
        ws.getCell('F66').value = d.glassH;
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I68').value = d.glassMark;
        ws.getCell('I70').value = d.glassType;
        ws.getCell('I72').value = d.glassManuf;
        ws.getCell('I75').value = d.overlap;
        ws.getCell('I76').value = d.sealing;

        // --- חישוב עומסי רוח (80-85) ---
        ws.getCell('I83').value = d.h_pressure;
        ws.getCell('I84').value = d.qb_val;
        ws.getCell('G82').value = d.we_val;
        ws.getCell('G83').value = d.fw_val;
        ws.getCell('G84').value = d.ceze_val;
        ws.getCell('G85').value = d.topRail_load;

        // --- טבלת תוצאות (107-117) ---
        const rows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'db': 114, 'e': 116 };
        for (let k in rows) {
            ws.getCell(`I${rows[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${rows[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${rows[k]}`).value = d[`stat_${k}`];
        }

        // --- מסקנות וחתימות (37-40) ---
        ws.getCell('L37').value = d.finalConclusion;
        ws.getCell('L38').value = d.inspectorName;
        ws.getCell('L39').value = d.approverName;
        ws.getCell('C40').value = d.comments;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Test_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error generating file"); }
});

app.listen(10000, '0.0.0.0');
