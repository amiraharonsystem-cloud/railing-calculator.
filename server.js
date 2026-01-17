const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static('public'));

const FILE_PATH = path.join(__dirname, 'reports_database.xlsx');

app.post('/add-row', async (req, res) => {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('בדיקות מעקה');
            // הגדרת עמודות לפי קובץ האקסל המקורי שלך
            worksheet.columns = [
                { header: 'תאריך בדיקה', key: 'testDate', width: 15 },
                { header: 'שם האתר', key: 'siteName', width: 25 },
                { header: 'קוד פרויקט', key: 'projectId', width: 15 },
                { header: 'שם המזמין', key: 'clientName', width: 25 },
                { header: 'מיקום בדיקה', key: 'location', width: 20 },
                { header: 'תיאור הפריט', key: 'itemDescription', width: 30 }
            ];
        }

        worksheet.addRow(data);
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).json({ message: 'הנתונים נוספו בהצלחה!' });
    } catch (error) {
        res.status(500).json({ message: 'שגיאה בשמירת הקובץ' });
    }
});

app.get('/download', (req, res) => {
    if (fs.existsSync(FILE_PATH)) res.download(FILE_PATH);
    else res.status(404).send('הקובץ לא קיים עדיין');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server is running on port ${PORT}`));
