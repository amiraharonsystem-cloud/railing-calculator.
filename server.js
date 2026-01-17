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

        // כותרת וזיהוי (שורות 3-10)
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; // ראשונה/חוזרת

        // לפני הבדיקה (V או X)
        ws.getCell('I14').value = d.planOk; 
        ws.getCell('I15').value = d.engOk;
        ws.getCell('I16').value = d.glassEngOk;

        // פרטי המערכת
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;

        // ערכים לחישוב (התאמה מלאה לקובץ שהעלית)
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L19').value = d.Fser; // עומס שירות
        ws.getCell('L20').value = d.P_load; // P = Fser(1/2 L2 + 1/2 L1)

        // תוצאות בדיקה (סעיף 10.3.4 ו-10.3.5)
        const mapping = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'db': 114, 'e': 116 };
        for (let key in mapping) {
            ws.getCell(`I${mapping[key]}`).value = d[`hor_${key}`];
            ws.getCell(`J${mapping[key]}`).value = d[`res_${key}`];
            ws.getCell(`L${mapping[key]}`).value = d[`stat_${key}`];
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) { res.status(500).send("שגיאה"); }
});

app.listen(10000, '0.0.0.0');
