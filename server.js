const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();

app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells } = req.body; 
        const workbook = new ExcelJS.Workbook();
        
        // טעינת הקובץ המקורי שלך כבסיס
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // לולאת הזרקה לכל תא ותא שמוגדר בממשק
        Object.keys(cells).forEach(address => {
            const value = cells[address];
            if (value !== undefined && value !== "") {
                const cell = ws.getCell(address);
                // המרה למספר במידת האפשר כדי לשמור על תקינות הנוסחאות באקסל
                cell.value = (isNaN(value) || value === "") ? value : Number(value);
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Full_Engineering_Report_1142.xlsx');
        
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (e) {
        console.error("Excel Error:", e);
        res.status(500).send("שגיאה ביצירת הקובץ");
    }
});

app.listen(3000, () => console.log('Server is live on http://localhost:3000'));
