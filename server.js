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

        // --- מיפוי תאים לפי הדו"ח שלך ---
        ws.getCell('L3').value = d.testDate;      // תאריך בדיקה
        ws.getCell('C4').value = d.siteName;      // שם האתר
        ws.getCell('L4').value = d.projectId;     // קוד פרויקט
        ws.getCell('C7').value = d.clientName;    // שם המזמין
        ws.getCell('L7').value = d.clientRep;     // נציג המזמין
        ws.getCell('C10').value = d.structure;    // תיאור המבנה
        ws.getCell('C11').value = d.itemDesc;     // תיאור הפריט
        ws.getCell('C12').value = d.location;     // מיקום (קומה/דירה)
        
        // --- מדידות הנדסיות ---
        ws.getCell('L22').value = parseFloat(d.L1) || 0; // מרחק L1
        ws.getCell('L24').value = parseFloat(d.L2) || 0; // מרחק L2
        ws.getCell('L26').value = d.L_fill || '-';       // מפתח לוח מליא L
        
        // --- עומסי רוח והערות ---
        ws.getCell('L30').value = d.windLoad || '-';     // חישוב עומס רוח
        ws.getCell('C40').value = d.comments || '';      // הערות בודק

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
