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

        // --- עמוד 1: כותרות וזיהוי (שורות 1-10) ---
        ws.getCell('H2').value = d.reportNum;
        ws.getCell('L2').value = d.replacedReportNum;
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; 

        // --- סעיף 2: תיאור המערכת והצהרות (שורות 11-16) ---
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;
        ws.getCell('D14').value = d.planOk;   // תכנון (V/X)
        ws.getCell('D15').value = d.engOk;    // חישוב שלד
        ws.getCell('D16').value = d.glassOk;  // חישוב מליא

        // --- סעיף 3: נתוני הבדיקה (שורות 19-32) ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;

        // --- חישובי רוח (10.3.5 ג') ---
        ws.getCell('I32').value = d.we;
        ws.getCell('K32').value = d.S_area;
        ws.getCell('L32').value = d.ws_total;

        // --- טבלת תוצאות מלאה (שורות 107-117) ---
        const map = { '107':'a', '108':'b', '110':'c', '112':'d', '114':'db', '116':'e' };
        Object.keys(map).forEach(row => {
            const key = map[row];
            ws.getCell(`I${row}`).value = d[`hor_${key}`];
            ws.getCell(`J${row}`).value = d[`res_${key}`];
            ws.getCell(`L${row}`).value = d[`stat_${key}`];
        });

        // --- חתימות (שורות 37-40) ---
        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.tester;
        ws.getCell('L39').value = d.approver;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send(e.message); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server on ${PORT}`));
