const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());

// פקודה שאומרת לשרת: חפש קבצים בתיקייה שבה אתה נמצא
const root = process.cwd();

app.get('/', (req, res) => {
    // הוא ינסה למצוא את הקובץ ב-public או בתיקייה הראשית
    const file1 = path.join(root, 'public', 'index.html');
    const file2 = path.join(root, 'index.html');

    if (fs.existsSync(file1)) return res.sendFile(file1);
    if (fs.existsSync(file2)) return res.sendFile(file2);
    
    res.status(404).send("File index.html not found. Check GitHub folder name.");
});

// הגשת שאר הקבצים (כמו script.js או css)
app.use(express.static(path.join(root, 'public')));
app.use(express.static(root));

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
