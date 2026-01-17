const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells } = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        Object.keys(cells).forEach(addr => {
            const val = cells[addr];
            if (val !== undefined && val !== "") {
                const cell = ws.getCell(addr);
                // הזרקה חכמה: מספרים כמספרים לשמירה על נוסחאות
                cell.value = (isNaN(val) || val.trim() === "") ? val : Number(val);
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(await workbook.xlsx.writeBuffer());
    } catch (e) {
        res.status(500).send("Error: " + e.message);
    }
});
app.listen(3000);
