javascript:(function(){
    let row = prompt("הדבק כאן את השורה שהעתקת מאקסל:");
    if(!row) return;
    let p = row.split(/\t/);
    let data = {
        order: p[0],
        project: p[3],
        address: p[4],
        date: new Date().toLocaleDateString('he-IL')
    };
    fetch('https://YOUR-APP-NAME.onrender.com/generate', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(data)
    })
    .then(response => response.blob())
    .then(blob => {
        let url = window.URL.createObjectURL(blob);
        let a = document.createElement('a');
        a.href = url;
        a.download = "Report.xlsx";
        a.click();
    });
})();
