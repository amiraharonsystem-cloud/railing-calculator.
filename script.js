const CITY_MAP = {"אבו גוש": 26, "אשדוד": 24, "ירושלים": 26, "תל אביב-יפו": 24 /* הוסף כאן את שאר הערים */};

// מילוי ערים בתיבת הבחירה
const citySelect = document.getElementById('city');
Object.keys(CITY_MAP).sort().forEach(city => {
    let opt = document.createElement('option');
    opt.value = city;
    opt.innerHTML = city;
    citySelect.appendChild(opt);
});

function toggleFields() {
    const isSteel = document.getElementById('type').value === 'פלדה';
    document.getElementById('windFields').classList.toggle('hidden', isSteel);
}

function calculate() {
    const type = document.getElementById('type').value;
    const l1 = parseFloat(document.getElementById('l1').value);
    const l2 = parseFloat(document.getElementById('l2').value);
    let Fser = 750;

    if (type === 'מילא') {
        const city = document.getElementById('city').value;
        const vb = CITY_MAP[city];
        const hb = parseFloat(document.getElementById('hb').value);
        // כאן מכניסים את הנוסחה המדויקת מהפייתון (Ce, We וכו')
        const ze = Math.max(hb + 1.05, 2.0);
        const ce = (0.234 * Math.log(ze / 1.0)) * (1 + 7 / (Math.log(ze / 1.0))); // דרגה IV כברירת מחדל
        const we = 0.613 * Math.pow(vb, 2) * ce * 1.8;
        const fw = we * 1.05;
        Fser = Math.ceil(Math.max(fw, 750 + (fw / 2)));
    }

    const res_b = Math.ceil(Fser * l1 * 0.75);
    const res_a = Math.ceil(res_b / 2);

    let output = `דוח פרויקט: ${document.getElementById('projId').value}\n`;
    output += `עומס נבחר (Fser): ${Fser} נ'/מ"א\n`;
    output += `סעיף א' (חצי עומס): ${res_a} נ'/מ"א\n`;
    output += `סעיף ב' (עומס מלא): ${res_b} נ'/מ"א\n`;

    const resDiv = document.getElementById('result');
    resDiv.innerText = output;
    resDiv.style.display = 'block';
    document.getElementById('downloadBtn').classList.remove('hidden');
}

function downloadPDF() {
    const text = document.getElementById('result').innerText;
    const blob = new Blob([text], { type: 'text/plain' });
    const anchor = document.createElement('a');
    anchor.download = `Report_${document.getElementById('city').value}.txt`;
    anchor.href = window.URL.createObjectURL(blob);
    anchor.click();
}