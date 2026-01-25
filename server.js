const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

app.post('/generate', async (req, res) => {
    try {
        const data = req.body;
        const workbook = new ExcelJS.Workbook();
        const templatePath = './template.xlsx';

        if (fs.existsSync(templatePath)) {
            await workbook.xlsx.readFile(templatePath);
        } else {
            workbook.addWorksheet('Report');
        }

        const worksheet = workbook.getWorksheet(1);

        // הזנת הנתונים לתאים
        worksheet.getCell('B2').value = data.date || "";
        worksheet.getCell('E2').value = "אמיר אהרון";
        worksheet.getCell('B3').value = data.project || "";
        worksheet.getCell('E3').value = data.order || "";
        worksheet.getCell('B4').value = data.address || "";

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        res.status(500).send(error.message);
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
