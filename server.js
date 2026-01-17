const express = require('express');
const ExcelJS = require('exceljs');
const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('template.xlsx'); // קובץ המקור שלך
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: פרטי זיהוי (שורות 1-10) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; // ראשונה/חוזרת

        // --- חלק 2: הצהרות ותיאור (שורות 11-16) ---
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;
        ws.getCell('I14').value = d.planStatus; // תכנון מעקה
        ws.getCell('I15').value = d.engStatus;  // חישוב שלד
        ws.getCell('I16').value = d.glassStatus; // חישוב לוח מליא

        // --- חלק 3: נתוני בדיקה וערכים לחישוב (סעיף 3) ---
        ws.getCell('L19').value = d.Fser;     // עומס שירות
        ws.getCell('L22').value = d.L1;       // מרחק בין ניצבים
        ws.getCell('L24').value = d.L2;       // מפתח שדה סמוך
        ws.getCell('L20').value = d.P_calc;   // עומס מחושב P
        ws.getCell('L26').value = d.L_fill;   // מפתח לוח מליא

        // --- חישוב עומס רוח (סעיף 10.3.5 ג') ---
        ws.getCell('I32').value = d.we_val;   // we (לחץ רוח)
        ws.getCell('K32').value = d.S_area;   // S (שטח הלוח)
        ws.getCell('L32').value = d.ws_total; // ws (עומס כולל)

        // --- טבלת תוצאות (תזוזה, שיורית, מסקנה) ---
        // מיפוי שורות לפי האקסל: 10.3.4א (107), 10.3.4ב (108), 10.3.5ג (110)
        const resMap = { 'a': 107, 'b': 108, 'c': 110 };
        for (let k in resMap) {
            ws.getCell(`I${resMap[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${resMap[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${resMap[k]}`).value = d[`stat_${k}`];
        }

        // --- חתימות וסיכום ---
        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.testerName;
        ws.getCell('L39').value = d.approverName;

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("ש
