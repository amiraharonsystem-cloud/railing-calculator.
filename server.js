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

        // --- חלק 1: זיהוי, לקוח ואתר (שורות 3-14) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.presenterName; // שם המציג
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structureDesc;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLocation;
        ws.getCell('C13').value = d.planningStatus;
        ws.getCell('C14').value = d.engStatus;

        // --- חלק 2: גאומטריה ועומסים (שורות 22-31) ---
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L30').value = d.ws_load;
        ws.getCell('L31').value = d.half_ws_load;

        // --- חלק 3: מפרט זכוכית ודיגוז (שורות 66-77) ---
        ws.getCell('B66').value = d.glassWidth;
        ws.getCell('E66').value = d.glassHeight;
        ws.getCell('H67').value = d.glassThick;
        ws.getCell('H68').value = d.glassMarking;
        ws.getCell('H70').value = d.glassType;
        ws.getCell('H72').value = d.glassBrand;
        ws.getCell('H75').value = d.overlap;
        ws.getCell('H76').value = d.sealing;

        // --- חלק 4: חישוב עומסי רוח (שורות 80-85) ---
        ws.getCell('C83').value = d.windHeight_h;
        ws.getCell('H82').value = d.qb_val;
        ws.getCell('G82').value = d.we_val;
        ws.getCell('G83').value = d.fw_val;
        ws.getCell('G84').value = d.ceze_val;
        ws.getCell('G85').value = d.topRailLoad;

        // --- חלק 5: טבלת תוצאות מפורטת (שורות 103-122) ---
        // 10.3.3
        ws.getCell('G35').value = d.res_1033;
        
        // 10.3.4 (א' עד ה')
        const rows1034 = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let key in rows1034) {
            let r = rows1034[key];
            ws.getCell(`I${r}`).value = d[`horiz_${key}`];
            ws.getCell(`J${r}`).value = d[`resid_${key}`];
            ws.getCell(`L${r}`).value = d[`status_${key}`];
        }

        // 10.3.5
        ws.getCell('L122').value = d.res_1035;

        // שדות גאומטריה חוזרים בתחתית הטבלה
        ws.getCell('L118').value = d.L1;
        ws.getCell('L120').value = d.L2;

        // --- חלק 6: סיכום (שורות 40-42) ---
        ws.getCell('C40').value = d.remarks;
        ws.getCell('C42').value = d.inspector;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Engineering_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("Error in generation"); }
});

app.listen(10000, '0.0.0.0');
