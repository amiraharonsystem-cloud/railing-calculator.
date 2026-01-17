const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const TEMPLATE_FILE = path.join(__dirname, 'template.xlsx');

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_FILE);
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: פרטי הדו"ח ---
        ws.getCell('L3').value = d.testDate;      
        ws.getCell('C4').value = d.siteName;      
        ws.getCell('L4').value = d.projectId;     

        // --- חלק 2: פרטי לקוח ומבנה ---
        ws.getCell('C7').value = d.clientName;    
        ws.getCell('L7').value = d.clientRep;     
        ws.getCell('C8').value = d.siteAddress;   
        ws.getCell('C10').value = d.structure;    
        ws.getCell('C11').value = d.itemDesc;     
        ws.getCell('C12').value = d.location;     
        ws.getCell('C13').value = d.planningStatus; 

        // --- חלק 3: מדידות L1, L2, L ---
        ws.getCell('L22').value = parseFloat(d.L1) || 0; 
        ws.getCell('L24').value = parseFloat(d.L2) || 0; 
        ws.getCell('L26').value = d.L_fill || '-';       

        // --- חלק 4: סעיפי בדיקה (V/X) ---
        ws.getCell('B35').value = d.check1033; // בדיקת עומס מרוכז
        ws.getCell('B36').value = d.check1034; // בדיקת עומס אופקי
        ws.getCell('B37').value = d.check1035; // בדיקת לוח מליא
        
        // --- חלק 5: עומס רוח והערות ---
        ws.getCell('L30').value = d.windLoad;
        ws.getCell('C40').value = d.comments;      
        ws.getCell('C42').value = d.inspectorName; 

        // שליחת הקובץ ישירות לדפדפן (פותר את בעיית השמירה ב-Render)
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        res.status(500).send("שגיאה ביצירת הקובץ: " + error.message);
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log(`Server is running`));
