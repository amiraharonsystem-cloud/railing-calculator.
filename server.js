const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());

// הגדרת נתיב לתיקיית public
const publicPath = path.join(__dirname, 'public');
app.use(express.static(publicPath));

// נתיב לקובץ האקסל
const FILE_PATH = path.join(__dirname, 'database.xlsx');

// פתרון שגיאת "Cannot GET /" - שליחת דף הבית למשתמש
app.get('/', (req, res) => {
    res.sendFile(path.join(publicPath, 'index.html'));
});

// הוספת נתונים לאקסל
app.post('/add-to-excel', async (req, res) => {
    const { name, email, phone, city } = req.body;
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('נתוני לקוחות');
            worksheet.columns = [
                { header: 'שם מלא', key: 'name', width: 20 },
                { header: 'אימייל', key: 'email', width: 25 },
                { header: 'טלפון', key: 'phone', width: 15 },
                { header: 'עיר', key: 'city', width: 15 }
            ];
        }

        worksheet.addRow({ name, email, phone, city });
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).json({ message: 'הנתונים נשמרו בהצלחה!' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'שגיאה בשמירת הנתונים' });
    }
});

// הורדת הקובץ
app.get('/download-excel', (req, res) => {
    if (fs.existsSync(FILE_PATH)) {
        res.download(FILE_PATH);
    } else {
        res.status(404).send('הקובץ עדיין לא נוצר. אנא הזן נתונים תחילה.');
    }
});

// הגדרת פורט עבור Render
const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});
