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
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // הזרקה דינמית לכל תא ותא שמופיע בטופס ה-HTML
        Object.keys(cells).forEach(address => {
            const value = cells[address];
            if (value !== undefined && value !== "") {
                const cell = ws.getCell(address);
                // אם הערך הוא מספר, נזין אותו כמספר כדי שנוסחאות האקסל יפעלו
                cell.value = isNaN(value) ? value : Number(value);
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Full_Railing_Report.xlsx');
        res.send(await workbook.xlsx.writeBuffer());
    } catch (e) {
        res.status(500).send("Error generating file: " + e.message);
    }
});

app.listen(3000, () => console.log('Server running on http://localhost:3000'));
