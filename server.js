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

        // --- חלק 1: זיהוי ואתר (שורות 3-14) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structureDesc;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLocation;
        ws.getCell('C13').value = d.planningStatus;
        ws.getCell('C14').value = d.engStatus;

        // --- חלק 2: גאומטריה ועומסי תכנון (שורות 22-31) ---
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L30').value = d.ws_load;
        ws.getCell('L31').value = d.half_ws_load;

        // --- חלק 3: סיכום, מסקנה וחתימות (שורות 37-41) ---
        ws.getCell('L37').value = d.finalStatus; // "עומד" / "לא עמד"
        ws.getCell('L38').value = d.inspectorName;
        ws.getCell('L39').value = d.approverName;
        ws.getCell('C40').value = d.finalRemarks;

        // --- חלק 4: מפרט זכוכית (שורות 66-77) ---
        ws.getCell('C66').value = d.glassWidth;
        ws.getCell('F66').value = d.glassHeight;
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I68').value = d.glassMarking;
        ws.getCell('I70').value = d.glassType;
        ws.getCell('I72').value = d.glassBrand;
        ws.getCell('I75').value = d.overlap;
        ws.getCell('I76').value = d.sealing;

        // --- חלק 5: חישוב עומסי רוח (שורות 82-85) ---
        ws.getCell('I83').value = d.h_val;   // גובה הפעלת לחץ
        ws.getCell('I84').value = d.qb_val;  // qb
        ws.getCell('G82').value = d.we_val;  // We
        ws.getCell('G83').value = d.fw_val;  // Fw
        ws.getCell('G84').value = d.ceze_val;// Ce(Ze)
        ws.getCell('G85').value = d.rail_load;// עומס על אזן עליון

        // --- חלק 6: טבלת תוצאות מפורטת (שורות 107-122) ---
        const testRows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let key in testRows) {
            let r = testRows[key];
            ws.getCell(`I${r}`).value = d[`horiz_${key}`]; // תזוזה אופקית
            ws.getCell(`J${r}`).value = d[`resid_${key}`]; // תזוזה שיורית
            ws.getCell(`L${r}`).value = d[`status_${key}`]; // סטטוס (עמד/לא עמד)
        }
        ws.getCell('L122').value = d.status_1035; // בדיקת מליא

        // תאי עזר חוזרים בתחתית
        ws.getCell('L118').value = d.L1;
        ws.getCell('L120').value = d.L2;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Engineering_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error generating file"); }
});

app.listen(10000, '0.0.0.0');
