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

// נתיב לקובץ האקסל בשרת
const FILE_PATH = path.join(__dirname, 'reports_database.xlsx');

// דף הבית - מציג את ה-HTML
app.get('/', (req, res) => {
    res.sendFile(path.join(publicPath, 'index.html'));
});

app.post('/add-row', async (req, res) => {
    try {
        const data = req.body;
        const workbook = new ExcelJS.Workbook();
        let worksheet;

        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('בדיקות מעקה');
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
        res.status(200).json({ message: 'הנתונים נשמרו בהצלחה באקסל!' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'שגיאה בשמירת הנתונים' });
    }
});

app.get('/download', (req, res) => {
    if (fs.existsSync(FILE_PATH)) {
        res.download(FILE_PATH);
    } else {
        res.status(404).send('הקובץ לא נוצר עדיין');
    }
});

// חשוב ל-Render: להאזין ל-0.0.0.0
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});
