
/* ===== טאבים ===== */
(function initTabs(){
  function activate(id){
    document.querySelectorAll('#tabs button').forEach(b=>b.classList.toggle('active', b.dataset.tab===id));
    document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active', t.id===id));
  }
  document.querySelectorAll('#tabs button').forEach(btn=>{
    btn.addEventListener('click', ()=> activate(btn.dataset.tab));
  });
})();

/* ===== נתונים ===== */
const CITY_MAP = {
  "אבו גוש": 26, "אופקים": 24, "אור יהודה": 24, "אור עקיבא": 24, "אילת": 24, "אלעד": 24,
  "אשדוד": 24, "אשקלון": 24, "באר שבע": 24, "בית שמש": 26, "בני ברק": 24, "בת ים": 24,
  "גבעתיים": 24, "הרצליה": 24, "חדרה": 24, "חולון": 24, "חיפה": 24, "טבריה": 24,
  "ירושלים": 26, "כפר סבא": 24, "מודיעין-מכבים-רעות": 26, "נתניה": 24, "פתח תקווה": 24,
  "צפת": 28, "ראשון לציון": 24, "רחובות": 24, "רמת גן": 24, "תל אביב-יפו": 24
};
const TERRAIN_DATA = {
  "דרגה 0: ים וחוף פתוח": { z0:0.003, kr:0.156 },
  "דרגה I: שטח מישורי ללא מכשולים": { z0:0.01,  kr:0.170 },
  "דרגה II: שטח פתוח עם מכשולים":   { z0:0.05,  kr:0.190 },
  "דרגה III: אזורי מגורים ותעשייה":  { z0:0.3,   kr:0.215 },
  "דרגה IV: שטחים עירוניים ובניינים":{ z0:1.0,   kr:0.234 }
};
let lastCalc = null;

/* ===== אתחול רשימות ===== */
(function initSelects(){
  const citySel = document.getElementById('city');
  Object.keys(CITY_MAP).sort().forEach(c => { const o=document.createElement('option'); o.textContent=c; citySel.appendChild(o); });
  const terrSel = document.getElementById('terrain');
  Object.keys(TERRAIN_DATA).forEach(k => { const o=document.createElement('option'); o.textContent=k; terrSel.appendChild(o); });
  terrSel.value = "דרגה IV: שטחים עירוניים ובניינים";
  document.getElementById('labelMapShow').textContent = JSON.stringify(LABEL_MAP, null, 2);
})();

/* ===== חישוב ===== */
function doCalc() {
  try {
    const railingType = document.getElementById('railingType').value;
    if (!railingType) { alert("אנא בחר סוג מעקה"); return; }
    const projId = (document.getElementById('projectId').value||"").trim();
    if (!projId) { alert("אנא הזן מספר פרויקט"); return; }

    const l1 = parseFloat(document.getElementById('l1').value);
    const l2 = parseFloat(document.getElementById('l2').value);
    const hr = parseFloat(document.getElementById('hr').value);

    let Fser, fw=0, we=0, vb=0, ce=0, hb=0;
    const unit = "נ׳/מ\"א", rle = "\u202b";

    if (railingType === "מעקה פלדה") {
      Fser = 750; fw = 0; we = 0;
    } else {
      const cityVal = document.getElementById('city').value;
      if (!cityVal) { alert("אנא בחר עיר"); return; }
      vb = CITY_MAP[cityVal] || 24;
      const terr = TERRAIN_DATA[document.getElementById('terrain').value];
      hb = parseFloat(document.getElementById('hb').value);
      const z0 = terr.z0, kr = terr.kr;
      const ze = Math.max(hb + hr, 2.0);
      const ln = Math.log(ze / z0);
      ce = (kr * ln) * (1 + 7 / ln);
      we = 0.613 * (vb**2) * ce * 1.8;
      fw = we * hr;
      Fser = Math.ceil(Math.max(fw, 750 + (fw/2)));
    }

    const res_b = Math.ceil(Fser * l1 * 0.75);
    const res_a = Math.ceil(res_b / 2);
    const res_g = Math.ceil(0.5 * Fser * (0.5 * l2 + 0.5 * l1));
    const res_d = Math.ceil(Fser * (0.5 * l2 + 0.125 * l1));
    const res_h = Math.ceil(Fser * (0.5 * l2 + 0.5 * l1));

    let res = "========================================================================\n";
    res += `               דוח תוצאות חישוב הנדסי - ${railingType}               \n`;
    res += "========================================================================\n\n";
    const dateStr = document.getElementById('inspectionDate').value || new Date().toLocaleDateString('he-IL');
    res += `${rle}תאריך: ${dateStr} | פרויקט: ${projId}\n`;
    if (railingType === "מעקה עם לוח מילא") {
      res += `${rle}מיקום: ${document.getElementById('city').value} | גובה בניין: ${document.getElementById('hb').value} מ׳\n`;
      res += `${rle}עומס רוח מחושב (Fw): ${fw.toFixed(2)} ${unit}\n`;
    } else {
      res += `${rle}סוג מעקה פלדה - חישוב על פי עומס שירות בסיסי 750 ${unit}\n`;
    }
    res += `${rle}גיאומטריה: מפתח א׳ = ${l1.toFixed(2)} מ׳, מפתח ב׳ = ${l2.toFixed(2)} מ׳\n`;
    res += `${rle}>> העומס הנבחר לחישוב (Fser): ${Fser} ${unit}\n\n`;
    res += `${rle}עומסים מרוכזים על ניצבים:\n`;
    res += "------------------------------------------------------------------------\n";
    res += `${rle}סעיף א׳ (חצי עומס): ${res_a} ${unit}\n`;
    res += `${rle}סעיף ב׳ (עומס מלא): ${res_b} ${unit}\n`;
    res += `${rle}סעיף ג׳ (ניצב ביניים): ${res_g} ${unit}\n`;
    res += `${rle}סעיף ד׳ (ניצב פינה): ${res_d} ${unit}\n`;
    res += `${rle}סעיף ה׳ (ניצב קצה): ${res_h} ${unit}\n`;

    document.getElementById('calcOutput').textContent = res;
    lastCalc = { railingType, projId, l1, l2, hr, fw, we, Fser, res_a, res_b, res_g, res_d, res_h,
                 dateStr, city: document.getElementById('city').value, hb, terrain: document.getElementById('terrain').value };

    const report = buildReportText(lastCalc);
    document.getElementById('reportArea').textContent = report;

    document.querySelector('#tabs button[data-tab="t4"]').click(); // מעבר לטאב תוצאות
  } catch (err) {
    alert("וודא שכל השדות מולאו כראוי: " + err);
  }
}
function buildReportText(c) {
  const u = 'נ׳/מ\"א', rle = "\u202b";
  let t = "תעודת בדיקת מעקה במעטפת הבניין\n==========================================\n";
  t += `${rle}תאריך בדיקה: ${c?.dateStr || ''}\n`;
  t += `${rle}שם האתר: ${document.getElementById('siteName').value}\n`;
  t += `${rle}קוד פרויקט: ${c?.projId || ''}\n\n`;
  t += "1. פרטי הלקוח וזיהוי ההזמנה:\n";
  t += `${rle}שם המזמין: ${document.getElementById('customerName').value}\n`;
  t += `${rle}מען המזמין: ${document.getElementById('customerAddress').value}\n`;
  t += `${rle}נציג המזמין בבדיקה: ${document.getElementById('customerRep').value}\n`;
  t += `${rle}כתובת האתר: ${document.getElementById('siteAddress').value}\n\n`;
  t += "2. פרטים על המערכת שנבדקה:\n";
  t += `${rle}תיאור המבנה: ${document.getElementById('structureDesc').value}\n`;
  t += `${rle}תיאור הפריט שנבדק: ${document.getElementById('itemDesc').value}\n`;
  t += `${rle}מיקום הבדיקה: ${document.getElementById('inspectionLocation').value}\n`;
  t += `${rle}סוג הבדיקה: בדיקה חזותית ובדיקה גיאומטרית\n\n`;
  t += "3. נתוני הבדיקה ותוצאות:\n";
  t += `${rle}Fser=${c.Fser} ${u} | L1=${c.l1.toFixed(2)}m | L2=${c.l2.toFixed(2)}m | hr=${c.hr.toFixed(2)}m\n`;
  if (c.railingType === "מעקה עם לוח מילא") {
    t += `${rle}עיר=${c.city} | hb=${c.hb}m | We=${c.we.toFixed(2)} נ׳/מ\"ר | Fw=${c.fw.toFixed(2)} ${u}\n`;
  } else {
    t += `${rle}מעקה פלדה: עומס שירות בסיסי 750 ${u}\n`;
  }
  t += `${rle}סעיף א׳=${c.res_a} | סעיף ב׳=${c.res_b} | סעיף ג׳=${c.res_g} | סעיף ד׳=${c.res_d} | סעיף ה׳=${c.res_h}\n\n`;
  t += "4. מסקנה:\n";
  t += `${rle}שם הבודק: ${document.getElementById('inspectorName').value}\n`;
  t += `${rle}שם המאשר ותפקידו: ${document.getElementById('approverTitle').value}\n`;
  t += `${rle}הערות: ${document.getElementById('notes').value}\n`;
  return t;
}

/* אירועים */
document.getElementById('btnCalc').addEventListener('click', doCalc);
document.getElementById('btnRecalc').addEventListener('click', doCalc);
document.getElementById('btnCopy').addEventListener('click', ()=> {
  const txt = document.getElementById('reportArea').textContent;
  navigator.clipboard.writeText(txt).then(()=>alert("הדוח הועתק ללוח"));
});

/* ===== מיפוי תוויות ===== */
const LABEL_MAP = [
  { label:"תאריך בדיקה:", value:()=>document.getElementById('inspectionDate').value || new Date().toLocaleDateString('he-IL'), offsetCol:1 },
  { label:"שם האתר:",     value:()=>document.getElementById('siteName').value, offsetCol:1 },
  { label:"קוד פרויקט:",  value:()=>document.getElementById('projectId').value, offsetCol:1 },
  { label:"חישוב עומס הרוח על האלמנט הנבדק Fw -", value:()=> (lastCalc ? `${lastCalc.fw.toFixed(2)} נ׳/מ\"א` : ""), offsetCol:1 },
  { label:"עומס רוח על אזן עליון (נ' למטר אורך)", value:()=> (lastCalc ? `${lastCalc.Fser} נ׳/מ\"א` : ""), offsetCol:1 },
  { label:"שם הבודק:",    value:()=>document.getElementById('inspectorName').value, offsetCol:1 },
  { label:"שם המאשר ותפקידו:", value:()=>document.getElementById('approverTitle').value, offsetCol:1 },
];
document.getElementById('labelMapShow').textContent = JSON.stringify(LABEL_MAP, null, 2);

/* ===== שתילה לאקסל ===== */
async function plantIntoExcel(){
  try {
    if (!lastCalc) { alert("בצע חישוב לפני שתילה בקובץ"); return; }
    const file = document.getElementById('excelTemplate').files?.[0];
    if (!file) { alert("אנא טען קובץ תבנית .xlsx"); return; }
    const outName = `תעודה_${lastCalc.projId}_${lastCalc.city || 'פלדה'}_${new Date().toLocaleDateString('he-IL')}.xlsx`;

    if (window.XlsxPopulate) {
      // --- XlsxPopulate ---
      const wb = await XlsxPopulate.fromDataAsync(file);
      const sheet0 = wb.sheet(0);
      LABEL_MAP.forEach(m => setNextToLabel(sheet0, m.label, m.value(), m.offsetCol));
      (document.getElementById('extraSheets').value || "")
        .split(',').map(s=>s.trim()).filter(Boolean)
        .forEach(name => {
          const sh = wb.addSheet(name);
          if (name.includes("קלט")) {
            sh.cell("A1").value("קלט");
            sh.cell("A3").value("סוג מעקה"); sh.cell("B3").value(lastCalc.railingType);
            sh.cell("A4").value("עיר"); sh.cell("B4").value(lastCalc.city);
            sh.cell("A5").value("גובה בניין hb"); sh.cell("B5").value(lastCalc.hb);
            sh.cell("A6").value("גובה מעקה hr"); sh.cell("B6").value(lastCalc.hr);
            sh.cell("A7").value("L1"); sh.cell("B7").value(lastCalc.l1);
            sh.cell("A8").value("L2"); sh.cell("B8").value(lastCalc.l2);
          } else if (name.includes("חישובים")) {
            sh.cell("A1").value("חישובים");
            sh.cell("A3").value("Fser"); sh.cell("B3").value(lastCalc.Fser);
            sh.cell("A4").value("We (נ׳/מ\"ר)"); sh.cell("B4").value(lastCalc.we.toFixed(2));
            sh.cell("A5").value("Fw (נ׳/מ\"א)"); sh.cell("B5").value(lastCalc.fw.toFixed(2));
            sh.cell("A7").value("סעיף א׳"); sh.cell("B7").value(lastCalc.res_a);
            sh.cell("A8").value("סעיף ב׳"); sh.cell("B8").value(lastCalc.res_b);
            sh.cell("A9").value("סעיף ג׳"); sh.cell("B9").value(lastCalc.res_g);
            sh.cell("A10").value("סעיף ד׳"); sh.cell("B10").value(lastCalc.res_d);
            sh.cell("A11").value("סעיף ה׳"); sh.cell("B11").value(lastCalc.res_h);
          }
        });
      const blob = await wb.outputAsync();
      const url = URL.createObjectURL(blob);
      downloadUrl(url, outName);

    } else if (window.XLSX) {
      // --- SheetJS (fallback) ---
      const ab = await file.arrayBuffer();
      const wb = XLSX.read(ab, { type: "array" });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];

      LABEL_MAP.forEach(m => sheetjsSetNextToLabel(wb, ws, m.label, m.value(), m.offsetCol || 1));

      (document.getElementById('extraSheets').value || "")
        .split(',').map(s=>s.trim()).filter(Boolean)
        .forEach(name => {
          const aoa = buildExtraSheetAOA(name);
          const nws = XLSX.utils.aoa_to_sheet(aoa);
          XLSX.utils.book_append_sheet(wb, nws, name);
        });

      XLSX.writeFile(wb, outName);
    } else {
      alert("לא נטענה ספריית Excel (XlsxPopulate/XLSX). בדוק את ה-CDN וההרשאות של הדפדפן.");
    }
  } catch (e) {
    alert("שגיאה בעת שתילה/הורדה: " + e);
  }
}
document.getElementById('btnPlant').addEventListener('click', plantIntoExcel);

/* עזרים */
function downloadUrl(url, filename){
  const a=document.createElement('a'); a.href=url; a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
function setNextToLabel(sheet, labelText, value, offsetCol) {
  const matches = sheet.find(labelText) || []; // XlsxPopulate Find API
  matches.forEach(cell => {
    const r = cell.rowNumber(); const c = cell.columnNumber();
    sheet.cell(r, c + (offsetCol||1)).value(value);
  });
}
function sheetjsSetNextToLabel(wb, ws, labelText, value, offsetCol){
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({r:R,c:C});
      const cell = ws[addr];
      const v = (cell && cell.v != null) ? String(cell.v).trim() : "";
      if (v === labelText) {
        const targetAddr = XLSX.utils.encode_cell({r:R, c:C + offsetCol});
        ws[targetAddr] = { t:'s', v:String(value) };
      }
    }
  }
}
function buildExtraSheetAOA(name){
  if (name.includes("קלט")) {
    return [
      ["קלט"],[],
      ["סוג מעקה", lastCalc.railingType],
      ["עיר", lastCalc.city],
      ["גובה בניין hb", lastCalc.hb],
      ["גובה מעקה hr", lastCalc.hr],
      ["L1", lastCalc.l1],
      ["L2", lastCalc.l2],
    ];
  }
  if (name.includes("חישובים")) {
    return [
      ["חישובים"],[],
      ["Fser", lastCalc.Fser],
      ["We (נ׳/מ\"ר)", Number(lastCalc.we.toFixed(2))],
      ["Fw (נ׳/מ\"א)", Number(lastCalc.fw.toFixed(2))],
      [],["סעיף א׳", lastCalc.res_a],
      ["סעיף ב׳", lastCalc.res_b],
      ["סעיף ג׳", lastCalc.res_g],
      ["סעיף ד׳", lastCalc.res_d],
      ["סעיף ה׳", lastCalc.res_h],
    ];
  }
  return [["גיליון חדש"]];
}
