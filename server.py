const express = require('express');
const axios = require('axios');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.static('public'));

const SHEET_ID = '1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs';
const GIDS = ['689162898', '1424845815'];

app.get('/api/schedule/:tester', async (req, res) => {
    const testerName = req.params.tester.split(' ')[0]; // מחפש לפי שם פרטי
    let foundProjects = [];

    try {
        for (let gid of GIDS) {
            const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
            const response = await axios.get(url);
            const rows = response.data.split('\n').map(r => r.split(',').map(c => c.replace(/"/g, '').trim()));

            rows.forEach((row, i) => {
                if (row.join(' ').includes(testerName)) {
                    // לוגיקה חכמה: אם השורה הנוכחית ריקה מפרטים, בדוק את השורה הבאה
                    let dataRow = (row[3] && row[3].length > 1) ? row : rows[i + 1];
                    
                    if (dataRow && dataRow[3] && !dataRow.join(' ').includes("בוטל")) {
                        foundProjects.push({
                            client: dataRow[3],
                            address: dataRow[4],
                            pNum: dataRow[1],
                            date: gid === '689162898' ? '2026-01-18' : '2026-01-19'
                        });
                    }
                }
            });
        }
        res.json(foundProjects);
    } catch (error) {
        res.status(500).json({ error: "Failed to fetch data" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
