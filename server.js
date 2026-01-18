const express = require('express');
const axios = require('axios');
const cors = require('cors');
const path = require('path');
const app = express();

app.use(cors());
app.use(express.static('public'));

const SHEET_ID = '1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs';
const GIDS = ['689162898', '1424845815']; // טאבים 18.01 ו-19.01

app.get('/api/schedule/:tester', async (req, res) => {
    try {
        const testerName = req.params.tester.split(' ')[0]; // חיפוש לפי שם פרטי (סרגיי/אמיר)
        let foundProjects = [];

        for (let gid of GIDS) {
            const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
            const response = await axios.get(url);
            const rows = response.data.split('\n').map(r => r.split(',').map(c => c.replace(/"/g, '').trim()));

            rows.forEach((row, i) => {
                if (row.join(' ').includes(testerName)) {
                    // לוגיקה לסרגיי: אם השורה שלו ריקה, קח נתונים מהשורה הבאה (i+1)
                    let dataRow = (row[3] && row[3].length > 1) ? row : rows[i + 1];
                    
                    if (dataRow && dataRow[3] && !dataRow.join(' ').includes("בוטל")) {
                        foundProjects.push({
                            client: dataRow[3],
                            address: dataRow[4] || dataRow[5] || "יהוד",
                            pNum: dataRow[1] || "148533",
                            date: gid === '689162898' ? '2026-01-18' : '2026-01-19'
                        });
                    }
                }
            });
        }
        res.json(foundProjects);
    } catch (error) {
        res.status(500).json({ error: "שגיאה במשיכת נתונים" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
