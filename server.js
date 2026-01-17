const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();

app.use(express.json({ limit: '50mb' })); // תמיכה בכמות נתונים גדולה מאוד
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells } = req.body;
        const workbook = new ExcelJS.Workbook();
        
        // טעינת הטמפלייט האמיתי שלך
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // הזרקה מוחלטת לכל תא ותא שנשלח מהווב
        Object.keys(cells).forEach(address => {
            const val = cells[address];
            if (val !== undefined && val !== "") {
                const cell = ws.getCell(address);
                // המרה למספר במידה והנתון הוא הנדסי (לשמירה על נוסחאות האקסל)
                cell.value = (isNaN(val) || val.trim() === "") ? val : Number(val);
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Engineering_Report_1142_Final.xlsx');
        
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (e) {
        console.error(e);
        res.status(500).send("Error generating file");
    }
});

app.listen(3000, () => console.log('Full System Running on http://localhost:3000'));
