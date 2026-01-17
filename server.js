const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: כותרת (שורות 1-6) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;

        // --- חלק 2: פרטי לקוח (שורות 7-20) ---
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.clientRep;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structureType;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.testLocation;
        ws.getCell('C13').value = d.planningStatus; // בחירה: הוגש/לא הוגש
        ws.getCell('C14').value = d.calcStatus;     // בחירה: הוגש/לא הוגש

        // --- חלק 3: מדידות טכניות (שורות 21-34) ---
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_gap;

        // --- חלק 4: עומסים וסעיפי תקן (שורות 35-130) ---
        ws.getCell('L30').value = d.ws_val;
        ws.getCell('L31').value = d.half_ws;
        ws.getCell('G35').value = d.res1033;
        ws.getCell('G36').value = d.res1034;
        ws.getCell('G37').value = d.res1035;

        // --- חלק 5: הערות סיום וחתימות (עד שורה 138) ---
        ws.getCell('C40').value = d.notes;
        ws.getCell('C42').value = d.testerName;
        // ניתן להוסיף כאן כל תא נוסף שקיים ב-138 השורות

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Railing_Report.xlsx`);
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        res.status(500).send("שגיאה: " + error.message);
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log('Live'));
