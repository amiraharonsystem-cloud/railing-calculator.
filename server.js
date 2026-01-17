const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());

// פתרון סופי לבעיית הנתיבים - בודק איפה אנחנו נמצאים
const rootDir = process.cwd();
const publicPath = path.join(rootDir, 'public');

// משרת קבצים סטטיים
app.use(express.static(publicPath));

const FILE_PATH = path.join(rootDir, 'maake_reports.xlsx');

// דף הבית - מנסה למצוא את הקובץ בכמה דרכים
app.get('/', (req, res) => {
    const locations = [
        path.join(publicPath, 'index.html'),
        path.join(rootDir, 'public', 'index.html'),
        path.join(__dirname, 'public', 'index.html')
    ];

    for (let loc of locations) {
        if (fs.existsSync(loc)) {
            return res.sendFile(loc);
        }
    }
    
    // אם הגענו לכאן, הקובץ באמת לא נמצא
    res.status(404).send(`Error: index.html not found. Locations searched: ${locations.join(', ')}`);
});

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
                { header: 'תאריך', key: 'date' },
                { header: 'אתר', key: 'site' },
                { header: 'פרויקט', key: 'code' },
                { header: 'מזמין', key: 'client' },
                { header: 'מיקום', key: 'loc' }
            ];
        }
        worksheet.addRow(req.body);
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).json({ message: 'נשמר בהצלחה!' });
    } catch (e) {
        res.status(500).json({ message: e.message });
    }
});

app.get('/download', (req, res) => {
    if (fs.existsSync(FILE_PATH)) res.download(FILE_PATH);
    else res.status(404).send('הקובץ לא קיים');
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});
