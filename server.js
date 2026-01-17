const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const TEMPLATE_FILE = path.join(__dirname, 'template.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'report_result.xlsx');

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_FILE);
        const ws = workbook.getWorksheet(1);

        // --- חלק 1: פרטי כלליים ותאריכים ---
        ws.getCell('L3').value = d.testDate;      
        ws.getCell('C4').value = d.siteName;      
        ws.getCell('L4').value = d.projectId;     
        
        // --- חלק 2: פרטי לקוח ---
        ws.getCell('C7').value = d.clientName;    
        ws.getCell('L7').value = d.clientRep;     
        ws.getCell('C8').value = d.siteAddress;   

        // --- חלק 3: תיאור המבנה והפריט ---
        ws.getCell('C10').value = d.structure;    
        ws.getCell('C11').value = d.itemDesc;     
        ws.getCell('C12').value = d.location;     
        ws.getCell('C13').value = d.planningStatus; // "הוגש / לא הוגש"

        // --- חלק 4: נתונים הנדסיים ומדידות ---
        ws.getCell('L22').value = parseFloat(d.L1) || 0; 
        ws.getCell('L24').value = parseFloat(d.L2) || 0; 
        ws.getCell('L26').value = d.L_fill || '-';       
        ws.getCell('L30').value = d.windLoad || '-';     

        // --- חלק 5: סיכום והערות ---
        ws.getCell('C40').value = d.comments;      
        ws.getCell('C42').value = d.inspectorName; 

        await workbook.xlsx.writeFile(OUTPUT_FILE);
        res.json({ success: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/download', (req, res) => {
    res.download(OUTPUT_FILE);
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log(`Server is running`));
