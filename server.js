const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();

app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        
        // טעינת הקובץ מהתיקייה המקומית
        const templatePath = path.join(__dirname, 'template.xlsx');
        await workbook.xlsx.readFile(templatePath);
        const ws = workbook.getWorksheet(1);

        // --- עמוד 1: זיהוי (שורות 1-16) ---
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; 
        
        // תיאור וסטטוס (C11-I16)
        ws.getCell('C11').value = d.structureType;
        ws.getCell('C12').value = d.itemDesc;
        ws.getCell('C13').value = d.testLocation;
        ws.getCell('I14').value = d.planStatus;
        ws.getCell('I15').value = d.engStatus;
        ws.getCell('I16').value = d.glassStatus;

        // --- עמוד 2: נתוני רוח הנדסיים (שורות 80-85) ---
        ws.getCell('I83').value = d.h_ground;
        ws.getCell('I84').value = d.qb;
        ws.getCell('G82').value = d.ce_z;
        ws.getCell('G83').value = d.cp_net;
        ws.getCell('G84').value = d.we_design;

        // --- עמוד 3: ערכים לחישוב (שורות 19-32) ---
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L26').value = d.L_fill;
        
        // חישוב עומס רוח סופי על הלוח
        ws.getCell('I32').value = d.we_final;
        ws.getCell('K32').value = d.S_area;
        ws.getCell('L32').value = d.ws_total;

        // --- טבלת תוצאות (שורות 107-116) ---
        const rows = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'e': 116 };
        for (let k in rows) {
            ws.getCell(`I${rows[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${rows[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${rows[k]}`).value = d[`stat_${k}`];
        }

        // --- סיכום (שורות 37-40) ---
        ws.getCell('L37').value = d.conclusion;
        ws.getCell('L38').value = d.tester;
        ws.getCell('L39').value = d.approver;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.end(await workbook.xlsx.writeBuffer());
        
    } catch (e) {
        console.error(e);
        res.status(500).send("שגיאה ביצירת הקובץ");
    }
});

// הגדרת הפורט בצורה שמתאימה ל-Render
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
