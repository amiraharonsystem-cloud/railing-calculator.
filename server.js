const express = require('express');
const axios = require('axios');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.static('public'));

const SHEET_ID = '1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs';
const GIDS = ['689162898', '1424845815'];

app.get('/api/schedule/:tester', async (req, res) => {
    try {
        const testerName = req.params.tester.split(' ')[0]; 
        let foundProjects = [];

        for (let gid of GIDS) {
            const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
            const response = await axios.get(url);
            
            // פירוק CSV שמתמודד עם פסיקים בתוך תאים (כמו בכתובות)
            const rows = response.data.split('\n').map(row => 
                row.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(cell => cell.replace(/"/g, '').trim())
            );

            for (let i = 0; i < rows.length; i++) {
                if (rows[i].join(' ').includes(testerName)) {
                    // תיקון לסרגיי: אם התא של שם המזמין (אינדקס 3) ריק, קח את השורה הבאה
                    let dataRow = (rows[i][3] && rows[i][3].length > 1) ? rows[i] : rows[i + 1];
                    
                    if (dataRow && dataRow[3] && !dataRow.join(' ').includes("בוטל")) {
                        foundProjects.push({
                            client: dataRow[3],
                            address: dataRow[4] || "יהוד",
                            pNum: dataRow[1] || "148533",
                            date: gid === '689162898' ? '18/01/2026' : '19/01/2026'
                        });
                    }
                }
            }
        }
        res.json(foundProjects);
    } catch (error) {
        res.status(500).json({ error: "Server error fetching data" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
