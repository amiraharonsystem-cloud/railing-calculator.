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
        const templatePath = path.join(__dirname, 'template.xlsx');
        
        await workbook.xlsx.readFile(templatePath);
        const ws = workbook.getWorksheet(1);

        // מיפוי שדות לפי האקסל שלך
        ws.getCell('L3').value = d.testDate;      
        ws.getCell('C4').value = d.siteName;      
        ws.getCell('L4').value = d.projectId;     
        ws.getCell('C7').value = d.clientName;    
        ws.getCell('L7').value = d.clientRep;     
        ws.getCell('C8').value = d.siteAddress;   
        ws.getCell('C10').value = d.structure;    
        ws.getCell('C11').value = d.itemDesc;     
        ws.getCell('C12').value = d.location;     
        ws.getCell('C13').value = d.planningStatus; 

        // מדידות טכניות
        ws.getCell('L22').value = d.L1; 
        ws.getCell('L24').value = d.L2; 
        ws.getCell('L26').value = d.L_fill; 
        ws.getCell('L30').value = d.windLoad;

        // תוצאות בדיקה (V/X)
        ws.getCell('B35').value = d.check1033;
        ws.getCell('B36').value = d.check1034;
        ws.getCell('B37').value = d.check1035;
        
        ws.getCell('C40').value = d.comments;      
        ws.getCell('C42').value = d.inspectorName; 

        // שליחה ישירה ללא שמירה בשרת
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Railing_Report.xlsx');
        
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send("שגיאה: וודא שהעלית קובץ template.xlsx לגיטהאב");
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log('Server Live'));
