const express = require('express');
const axios = require('axios');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.static('public'));

const SHEET_ID = '1JpM_HhT-EyTclnSP12VlpaZFKZwgfHSGOS4YRPZbbvs';
const GIDS = ['689162898', '1424845815']; // 18.01 and 19.01

app.get('/api/schedule/:tester', async (req, res) => {
    try {
        const testerSearch = decodeURIComponent(req.params.tester).split(' ')[0];
        let foundProjects = [];

        for (let gid of GIDS) {
            const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
            const response = await axios.get(url, { timeout: 10000 });
            
            // פירוק ה-CSV למערך תוך שמירה על תאים עם פסיקים
            const rows = response.data.split('\n').map(row => 
                row.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(cell => cell.replace(/"/g, '').trim())
            );

            rows.forEach((row, i) => {
                const fullRowText = row.join(' ');
                if (fullRowText.includes(testerSearch)) {
                    // לוגיקת סרגיי: אם השורה הנוכחית לא מכילה שם פרויקט, בדוק את השורה הבאה
                    let dataRow = (row[3] && row[3].length > 2) ? row : (rows[i + 1] || row);
                    
                    if (dataRow[3] && !dataRow.join(' ').includes("בוטל")) {
                        foundProjects.push({
                            client: dataRow[3],
                            address: dataRow[4] || "יהוד",
                            pNum: dataRow[1] || "148533",
                            date: gid === '689162898' ? '18/01/2026' : '19/01/2026'
                        });
                    }
                }
            });
        }
        res.json(foundProjects);
    } catch (error) {
        console.error("Server Error:", error.message);
        res.status(500).json({ error: "שגיאת שרת פנימית" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
