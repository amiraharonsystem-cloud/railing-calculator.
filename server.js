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

        // --- כותרת ופרטי אתר ---
        ws.getCell('L3').value = d.testDate;      
        ws.getCell('C4').value = d.siteName;      
        ws.getCell('L4').value = d.projectId;     

        // --- 1. פרטי לקוח וזיהוי (שורות 7-14) ---
        ws.getCell('C7').value = d.clientName;    
        ws.getCell('L7').value = d.clientRep;     
        ws.getCell('C8').value = d.clientAddress; 
        ws.getCell('C10').value = d.structureType;
        ws.getCell('C11').value = d.itemDescription;
        ws.getCell('C12').value = d.testLocation;
        ws.getCell('C13').value = d.planningStatus; // תכנון המעקה
        ws.getCell('C14').value = d.calcStatus;     // חישוב הנדסי

        // --- 2. נתוני גיאומטריה (שורות 22-26) ---
        ws.getCell('L22').value = d.L1; // מרחק בין ניצבים
        ws.getCell('L24').value = d.L2; // מפתח שדה סמוך
        ws.getCell('L26').value = d.L_gap; // מפתח לוח מליא

        // --- 3. עומסי רוח (שורות 30-31) ---
        ws.getCell('L30').value = d.ws_load; // we x S
        ws.getCell('L31').value = d.half_ws; // 0.5ws

        // --- 4. טבלת תוצאות (שורות 35-37) ---
        ws.getCell('G35').value = d.res1033; 
        ws.getCell('G36').value = d.res1034;
        ws.getCell('G37').value = d.res1035;

        // --- 5. הערות וסיכום ---
        ws.getCell('C40').value = d.comments;
        ws.getCell('C42').value = d.inspectorName;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Report_${d.projectId}.xlsx`);
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        res.status(500).send("שגיאה: וודא שקובץ template.xlsx נמצא בגיטהאב");
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log('Server Running'));
