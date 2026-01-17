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

        // עמוד 1 - פרטי המזמין והאתר
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress; // מען המזמין
        ws.getCell('L7').value = d.presenterName; // נציג המזמין
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; // ראשונה / חוזרת

        // סטטוס לפני בדיקה (V/X)
        ws.getCell('I14').value = d.planOk;
        ws.getCell('I15').value = d.engOk;
        ws.getCell('I16').value = d.glassOk;

        // תיאור המערכת
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;

        // סעיף 3: נתונים לחישוב (הלב של הדוח)
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;

        // עומס רוח (סעיף 10.3.5 ג')
        ws.getCell('I32').value = d.We;
        ws.getCell('K32').value = d.S_area;
        ws.getCell('L32').value = d.Ws_total;

        // טבלת תוצאות (I=תזוזה, J=שיורית, L=מסקנה)
        const mapping = { '1034a': 107, '1034b': 108, '1035c': 110 };
        Object.keys(mapping).forEach(key => {
            ws.getCell(`I${mapping[key]}`).value = d[`hor_${key}`];
            ws.getCell(`J${mapping[key]}`).value = d[`res_${key}`];
            ws.getCell(`L${mapping[key]}`).value = d[`stat_${key}`];
        });

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("שגיאה"); }
});
app.listen(10000);
