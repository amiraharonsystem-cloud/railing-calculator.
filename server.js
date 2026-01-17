const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());

// מציאת הנתיב הנוכחי של השרת
const currentDir = process.cwd();

// פונקציה חכמה למציאת index.html
app.get('/', (req, res) => {
    const pathsToTry = [
        path.join(currentDir, 'public', 'index.html'),
        path.join(currentDir, 'index.html'),
        path.join(__dirname, 'public', 'index.html')
    ];

    for (let p of pathsToTry) {
        if (fs.existsSync(p)) {
            return res.sendFile(p);
        }
    }
    res.status(404).send("Error: index.html not found. Please check Root Directory in Render Settings.");
});

// הגדרת תיקיית קבצים סטטיים
app.use(express.static(path.join(currentDir, 'public')));

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
