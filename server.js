const express = require('express');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
app.use(bodyParser.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const { cells } = req.body;
        const workbook = new ExcelJS.Workbook();
        
        // טעינת תבנית האקסל המקורית
        await workbook.xlsx.readFile(path.join(__dirname, 'template.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        // הזרקת כל השדות לפי הכתובות שנשלחו מהממשק
        // זה כולל את שורות 2-14, 53-64 (גאומטריה), 100-117 (עומסים) ו-123-145 (מליא וסיכום)
        Object.keys(cells).forEach(addr => {
            const cell = worksheet.getCell(addr);
            cell.value = cells[addr];
        });

        // הגדרת כותרות התגובה להורדת קובץ
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=report.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error generating excel:', error);
        res.status(500).send('שגיאה ביצירת הקובץ');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
