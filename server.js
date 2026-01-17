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
const OUTPUT_FILE = path.join(__dirname, 'output_report.xlsx');

app.post('/generate-excel', async (req, res) => {
    try {
        const data = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_FILE);
        const worksheet = workbook.getWorksheet(1);

        // --- שיבוץ נתונים לפי הקובץ שלך ---
        worksheet.getCell('L3').value = data.testDate;      // תאריך בדיקה
        worksheet.getCell('C4').value = data.siteName;      // שם האתר
        worksheet.getCell('L4').value = data.projectId;     // קוד פרויקט
        worksheet.getCell('C7').value = data.clientName;    // שם המזמין
        worksheet.getCell('C11').value = data.location;     // מיקום (קומה/דירה)
        
        // שדות מדידה טכניים (מהסניפט שלך)
        worksheet.getCell('L22').value = parseFloat(data.L1); // מרחק בין ניצבים L1
        worksheet.getCell('L24').value = parseFloat(data.L2); // מפתח שדה סמוך L2

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
app.listen(PORT, '0.0.0.0', () => console.log(`Server running`));
