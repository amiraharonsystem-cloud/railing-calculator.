const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells } = req.body; // מקבל אובייקט שבו המפתח הוא שם התא, למשל {"H2": "123"}
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const ws = workbook.getWorksheet(1);

        // לולאה שעוברת על כל תא ותא שנשלח מהטופס ומזינה אותו לאקסל
        Object.keys(cells).forEach(cellAddress => {
            if (cells[cellAddress]) {
                ws.getCell(cellAddress).value = cells[cellAddress];
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Report_1142.xlsx');
        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { 
        console.error(e);
        res.status(500).send(e.message); 
    }
});

app.listen(3000, () => console.log('Server running on port 3000'));
