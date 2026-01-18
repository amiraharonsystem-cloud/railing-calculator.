const express = require('express');
const axios = require('axios');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.static('public'));

const SHEET_ID = '1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs';
const TAB_GIDS = ['689162898', '1424845815']; 

app.get('/api/schedule/:tester', async (req, res) => {
    try {
        const testerName = decodeURIComponent(req.params.tester).split(' ')[0];
        let foundProjects = [];

        for (let gid of TAB_GIDS) {
            const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
            const response = await axios.get(url);
            
            const rows = response.data.split('\n').map(row => 
                row.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(cell => cell.replace(/"/g, '').trim())
            );

            for (let i = 0; i < rows.length; i++) {
                if (rows[i].join(' ').includes(testerName)) {
                    // אם השורה של הבודק ריקה (כמו אצל סרגיי), קח את השורה הבאה
                    let dataRow = (rows[i][3] && rows[i][3].length > 2) ? rows[i] : (rows[i+1] || rows[i]);
                    
                    if (dataRow[3] && !dataRow.join(' ').includes("בוטל")) {
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
        res.status(500).json({ error: "Server Error" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running`));
