const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());

// הגדרת נתיב לתיקייה הציבורית
const publicPath = path.resolve(__dirname, 'public');
app.use(express.static(publicPath));

const FILE_PATH = path.join(__dirname, 'maake_reports.xlsx');

// דף הבית - מגיש את הקובץ מהתיקייה
app.get('/', (req, res) => {
    const indexPath = path.join(publicPath, 'index.html');
    if (fs.existsSync(indexPath)) {
        res.sendFile(indexPath);
    } else {
        res.status(404).send("שגיאה: קובץ index.html לא נמצא בתיקיית public");
    }
});

// הוספת שורה לאקסל
app.post('/add-row', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        let worksheet;
        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('בדיקות');
            worksheet.columns = [
                { header: 'תאריך', key: 'date', width: 15 },
                { header: 'שם אתר', key: 'site', width: 25 },
                { header: 'קוד פרויקט', key: 'code', width: 15 },
                { header: 'שם מזמין', key: 'client', width: 20 },
                { header: 'מיקום', key: 'loc', width: 20 }
            ];
        }
        worksheet.addRow(req.body);
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).json({ message: 'הנתונים נשמרו בהצלחה!' });
    } catch (e) {
        res.status(500).json({ message: e.message });
    }
});

// הורדת הקובץ
app.get('/download', (req, res) => {
    if (fs.existsSync(FILE_PATH)) {
        res.download(FILE_PATH);
    } else {
        res.status(404).send('הקובץ עדיין לא נוצר');
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});
