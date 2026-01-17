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
        
        // טעינת קובץ המקור (Template)
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // הזרקת הנתונים לתאים המדויקים
        Object.keys(cells).forEach(address => {
            const value = cells[address];
            if (value !== undefined && value !== "") {
                const cell = ws.getCell(address);
                // בדיקה אם הערך הוא מספרי כדי לשמור על תקינות חישובים באקסל
                if (!isNaN(value) && value.trim() !== "") {
                    cell.value = Number(value);
                } else {
                    cell.value = value;
                }
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Engineering_Report_1142.xlsx');
        
        const buffer = await workbook.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        console.error("Excel Error:", error);
        res.status(500).send("Error generating file");
    }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`System running on http://localhost:${PORT}`));
