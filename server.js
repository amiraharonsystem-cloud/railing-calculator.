const express = require('express');
const ExcelJS = require('exceljs');
const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
    try {
        const d = req.body;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('template.xlsx');
        const ws = workbook.getWorksheet(1);

        // עמוד 1: פרטי המזמין והצהרות (שורות 3-16)
        ws.getCell('L3').value = d.testDate;
        ws.getCell('L4').value = d.projectId;
        ws.getCell('C4').value = d.siteName;
        ws.getCell('C7').value = d.clientName;
        ws.getCell('I7').value = d.clientAddress;
        ws.getCell('L7').value = d.presenterName;
        ws.getCell('C8').value = d.siteAddress;
        ws.getCell('D9').value = d.testType; 
        ws.getCell('I14').value = d.planOk;
        ws.getCell('I15').value = d.engOk;
        ws.getCell('I16').value = d.glassOk;

        // עמוד 2: מפרט זכוכית וחישובי רוח (שורות 66-85)
        ws.getCell('C66').value = d.glassW;
        ws.getCell('F66').value = d.glassH;
        ws.getCell('I67').value = d.glassThick;
        ws.getCell('I68').value = d.glassMark;
        ws.getCell('I70').value = d.glassType;
        ws.getCell('I72').value = d.glassManuf;
        ws.getCell('I75').value = d.overlap;
        ws.getCell('I76').value = d.sealing;
        
        // נתוני רוח הנדסיים
        ws.getCell('I83').value = d.h_pressure;
        ws.getCell('I84').value = d.qb_val;
        ws.getCell('G82').value = d.we_val;
        ws.getCell('G83').value = d.fw_val;
        ws.getCell('G84').value = d.ceze_val;

        // עמוד 3: נתונים לחישוב ותוצאות (שורות 19-32 ו-107-117)
        ws.getCell('L19').value = d.Fser;
        ws.getCell('L20').value = d.P_calc;
        ws.getCell('L22').value = d.L1;
        ws.getCell('L24').value = d.L2;
        ws.getCell('L26').value = d.L_fill;
        ws.getCell('L30').value = d.We_calc; 
        ws.getCell('L31').value = d.S_area;
        ws.getCell('L32').value = d.Ws_total;

        // טבלת תוצאות סופית
        const map = { 'a': 107, 'b': 108, 'c': 110, 'd': 112, 'db': 114, 'e': 116 };
        for (let k in map) {
            ws.getCell(`I${map[k]}`).value = d[`hor_${k}`];
            ws.getCell(`J${map[k]}`).value = d[`res_${k}`];
            ws.getCell(`L${map[k]}`).value = d[`stat_${k}`];
        }

        res.end(await workbook.xlsx.writeBuffer());
    } catch (e) { res.status(500).send("Error"); }
});
app.listen(10000);
