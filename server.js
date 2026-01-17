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
        // טעינת הטמפלייט המקורי שלך
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // הזרקה אבסולוטית של כל תא ותא שנשלח מהממשק
        Object.keys(cells).forEach(address => {
            const cell = ws.getCell(address);
            cell.value = cells[address];
            // שמירה על פורמט בסיסי אם מדובר במספר
            if (!isNaN(cells[address]) && cells[address] !== "") {
                cell.value = Number(cells[address]);
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Full_Report_1142.xlsx');
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (e) {
        res.status(500).send("Error: " + e.message);
    }
});

app.listen(3000, () => console.log('Server running on http://localhost:3000'));
