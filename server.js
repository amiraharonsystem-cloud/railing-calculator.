const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static('public'));

const FILE_PATH = path.join(__dirname, 'database.xlsx');

app.post('/add-to-excel', async (req, res) => {
    const { name, email, phone, city } = req.body;
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        if (fs.existsSync(FILE_PATH)) {
            await workbook.xlsx.readFile(FILE_PATH);
            worksheet = workbook.getWorksheet(1);
        } else {
            worksheet = workbook.addWorksheet('Data');
            worksheet.columns = [
                { header: 'שם מלא', key: 'name', width: 20 },
                { header: 'אימייל', key: 'email', width: 25 },
                { header: 'טלפון', key: 'phone', width: 15 },
                { header: 'עיר', key: 'city', width: 15 }
            ];
        }
        worksheet.addRow({ name, email, phone, city });
        await workbook.xlsx.writeFile(FILE_PATH);
        res.status(200).json({ message: 'נשמר בהצלחה!' });
    } catch (e) { res.status(500).send('Error'); }
});

app.get('/download-excel', (req, res) => {
    if (fs.existsSync(FILE_PATH)) res.download(FILE_PATH);
    else res.status(404).send('No file yet');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server on port ${PORT}`));
