const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();

app.use(bodyParser.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells, fileName } = req.body;
        const workbook = new ExcelJS.Workbook();
        
        // טעינת התבנית הקיימת מהשרת
        const templatePath = path.join(__dirname, 'template.xlsx');
        await workbook.xlsx.readFile(templatePath);
        
        const worksheet = workbook.getWorksheet(1); // עובד על הגיליון הראשון

        // לולאה שעוברת על כל הכתובות שנשלחו ושותלת אותן באקסל
        Object.keys(cells).forEach(addr => {
            const cell = worksheet.getCell(addr);
            cell.value = cells[addr];
            
            // שמירה על העיצוב המקורי של התא (גבולות ופונט)
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(fileName)}`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error generating Excel:', error);
        res.status(500).send('שגיאה ביצירת הקובץ');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
