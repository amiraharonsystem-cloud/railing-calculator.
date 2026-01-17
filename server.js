const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();

app.get('/', (req, res) => {
    // השרת יבדוק ב-4 מקומות שונים עד שימצא את הקובץ
    const locations = [
        path.join(__dirname, 'index.html'),
        path.join(__dirname, 'public', 'index.html'),
        path.join(process.cwd(), 'index.html'),
        path.join(process.cwd(), 'public', 'index.html')
    ];

    for (let loc of locations) {
        if (fs.existsSync(loc)) {
            return res.sendFile(loc);
        }
    }
    res.status(404).send("Last check: File is missing from GitHub. Please upload index.html directly to the main folder.");
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is live on port ${PORT}`);
});
