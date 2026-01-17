const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const FILE_PATH = path.join(__dirname, 'maake_reports.xlsx');

app.post('/add-row', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('בדיקות מעקה');
            worksheet.columns = [
                { header: 'תאריך בדיקה', key: 'date', width: 15 },
                { header: 'שם האתר', key: 'site', width: 25 },
                { header: 'קוד פרויקט', key: 'code', width: 15 },
                { header: 'שם המזמין', key: 'client', width: 25 },
                { header: 'מיקום בדיקה', key: 'loc', width: 20 }
            ];
        }
        worksheet.addRow(req.body);
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).send({ message: 'נשמר בהצלחה!' });
    } catch (e) { res.status(500).send(e.message); }
});

app.get('/download', (req, res) => {
    if (fs.existsSync(FILE_PATH)) res.download(FILE_PATH);
    else res.status(404).send('אין קובץ');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`Running on ${PORT}`));
