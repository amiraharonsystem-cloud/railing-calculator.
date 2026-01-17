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

        // --- סעיף 1: פרטי המזמין והאתר ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('L7').value = d.clientRep;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('C10').value = d.structureDesc;
        ws.getCell('C11').value = d.itemDesc;
        ws.getCell('C12').value = d.location;
        ws.getCell('C13').value = d.planningStatus; 
        ws.getCell('C14').value = d.calcStatus;

        // --- סעיף 2: נתונים הנדסיים (מדידות) ---
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;

        // --- סעיף 3: עומסים ובדיקות ---
        ws.getCell('L30').value = d.ws_load;
        ws.getCell('L31').value = d.half_ws;
        ws.getCell('G35').value = d.res1033;
        ws.getCell('G36').value = d.res1034;
        ws.getCell('G37').value = d.res1035;

        // --- סיכום ---
        ws.getCell('C40').value = d.comments;
        ws.getCell('C42').value = d.inspectorName;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Report.xlsx`);
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        res.status(500).send("שגיאה: וודא שקובץ template.xlsx נמצא בשרת");
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log('Ready'));
