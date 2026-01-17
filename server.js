const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// שם קובץ התבנית שהעלית (וודא שהעלית אותו ל-GitHub באותו שם)
const TEMPLATE_FILE = path.join(__dirname, 'template.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'report_output.xlsx');

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/generate-report', async (req, res) => {
    try {
        const data = req.body;
        const workbook = new ExcelJS.Workbook();
        
        // טעינת התבנית המקורית שלך
        await workbook.xlsx.readFile(TEMPLATE_FILE);
        const worksheet = workbook.getWorksheet(1);

        // שיבוץ הנתונים בתאים המדויקים לפי הקובץ שלך
        worksheet.getCell('L3').value = data.testDate;      // תאריך בדיקה
        worksheet.getCell('C4').value = data.siteName;      // שם האתר
        worksheet.getCell('L4').value = data.projectId;     // קוד פרויקט
        worksheet.getCell('C7').value = data.clientName;    // שם המזמין
        worksheet.getCell('C11').value = data.location;     // מיקום הבדיקה
        worksheet.getCell('C10').value = data.itemDesc;     // תיאור הפריט

        // כאן ניתן להוסיף עוד שדות לפי הצורך (עומסים, מידות וכו')
        // הנוסחאות הקיימות באקסל יתעדכנו אוטומטית כשתפתח את הקובץ

        await workbook.xlsx.writeFile(OUTPUT_FILE);
        res.status(200).json({ message: 'הדו"ח הופק בהצלחה!' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'שגיאה ביצירת הקובץ: ' + error.message });
    }
});

app.get('/download', (req, res) => {
    if (fs.existsSync(OUTPUT_FILE)) res.download(OUTPUT_FILE);
    else res.status(404).send('הקובץ לא נמצא');
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => console.log(`Server running on port ${PORT}`));
