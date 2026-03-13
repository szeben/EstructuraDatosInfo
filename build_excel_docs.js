
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const workbook = XLSX.readFile('diccionario_datos_spi.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const tables = {};
let currentTableName = '';

// Skipping header rows (0, 1) based on inspection
for (let i = 2; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;

    const maybeTableName = row[0] ? String(row[0]).trim() : '';
    const seq = row[1];
    const colName = row[2] ? String(row[2]).trim() : '';
    const colType = row[3] ? String(row[3]).trim() : '';
    const colLen = row[4];
    const colDec = row[5];
    const colDesc = row[6] ? String(row[6]).trim() : '';

    if (maybeTableName && !seq) {
        // This is a header row for a new table or just noise
        // But in the sample, the table name is on the same row as the first column usually?
        // Wait, looking at row 2: "ACT_RET", 1, "CIA_CODCIA"...
        // So if maybeTableName exists and seq is present, it's the first row of a table.
    }

    if (maybeTableName) {
        currentTableName = maybeTableName;
        if (!tables[currentTableName]) {
            tables[currentTableName] = { columns: [] };
        }
    }

    if (colName && currentTableName) {
        tables[currentTableName].columns.push({
            name: colName,
            type: colType,
            len: colLen,
            dec: colDec,
            desc: colDesc
        });
    }
}

const outputDir = 'b:/Antigravity/BD_Infocent/docs';
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

function getModule(tname) {
    if (tname.startsWith('NM')) return { name: 'Nómina', icon: '📊' };
    if (tname.startsWith('SPI') || tname.startsWith('SS')) return { name: 'Seguridad', icon: '🛡️' };
    if (tname.startsWith('EO')) return { name: 'Estructura', icon: '🏢' };
    if (tname.startsWith('AR')) return { name: 'Configuración', icon: '⚙️' };
    if (tname.startsWith('RH')) return { name: 'RRHH', icon: '📜' };
    if (tname.startsWith('TA')) return { name: 'Maestras', icon: '📋' };
    if (tname.startsWith('RS')) return { name: 'Reclutamiento', icon: '👥' };
    return { name: 'Otros', icon: '📦' };
}

const tableNames = Object.keys(tables).sort();
let tableCards = "";

for (const tname of tableNames) {
    const info = tables[tname];
    const mod = getModule(tname);
    
    let rows = info.columns.map(c => `
        <tr>
            <td><span class="col-name">${c.name}</span></td>
            <td>${c.type}</td>
            <td>${c.len || ''}${c.dec ? ','+c.dec : ''}</td>
            <td class="desc-cell">${c.desc || ''}</td>
        </tr>
    `).join("");

    tableCards += `
    <div class="table-card" id="${tname}" data-name="${tname}">
        <div class="table-header-flex">
            <h3>${mod.icon} ${tname}</h3>
            <span class="table-tag">${mod.name}</span>
        </div>
        <table class="data-table">
            <thead>
                <tr><th>Campo</th><th>Tipo</th><th>Long</th><th>Descripción</th></tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>
    </div>
    `;
}

const css = `
:root {
    --primary: #6366f1;
    --bg-dark: #0f172a;
    --card-bg: #1e293b;
    --text-main: #f8fafc;
    --text-muted: #94a3b8;
    --accent: #38bdf8;
    --border: rgba(255, 255, 255, 0.1);
}
body { font-family: 'Inter', sans-serif; background: var(--bg-dark); color: var(--text-main); font-size: 0.85rem; margin: 0; padding: 0; }
.container { max-width: 1400px; margin: 0 auto; padding: 1rem; }
header { padding: 2rem; border-bottom: 1px solid var(--border); text-align: center; }
h1 { font-size: 1.8rem; background: linear-gradient(to right, #6366f1, #38bdf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
nav { display: flex; gap: 1rem; justify-content: center; margin: 1rem 0; }
nav a { color: white; text-decoration: none; padding: 0.3rem 0.8rem; border-radius: 4px; background: rgba(255,255,255,0.05); }
.search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 1rem 0; z-index: 1000; border-bottom: 1px solid var(--border); }
.search-box { width: 100%; padding: 0.6rem; border-radius: 6px; border: 1px solid var(--border); background: var(--card-bg); color: white; }
.table-card { background: var(--card-bg); border-radius: 8px; padding: 1rem; margin-bottom: 2rem; border: 1px solid var(--border); }
.table-header-flex { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem; }
.data-table { width: 100%; border-collapse: collapse; font-size: 0.75rem; }
.data-table th { text-align: left; padding: 0.5rem; background: rgba(99, 102, 241, 0.1); color: var(--accent); border-bottom: 1px solid var(--border); }
.data-table td { padding: 0.4rem; border-bottom: 1px solid var(--border); }
.col-name { color: #facc15; font-weight: 600; }
.table-tag { font-size: 0.65rem; background: var(--primary); padding: 2px 6px; border-radius: 10px; }
.desc-cell { color: var(--text-muted); font-style: italic; }
`;

const html = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Diccionario Excel - Infocent</title>
    <style>${css}</style>
</head>
<body>
    <header>
        <h1>Diccionario Técnico Base de Datos</h1>
        <p style="color: var(--text-muted);">${tableNames.length} Tablas importadas del Diccionario Excel</p>
    </header>
    <div class="container">
        <nav><a href="index.html">Resumen</a><a href="relaciones.html">Relaciones</a></nav>
        <div class="search-container">
            <input type="text" id="masterSearch" class="search-box" placeholder="Filtrar tablas o campos..." onkeyup="filter()">
        </div>
        <div id="tableList">${tableCards}</div>
    </div>
    <script>
        function filter() {
            let q = document.getElementById('masterSearch').value.toUpperCase();
            let cards = document.getElementsByClassName('table-card');
            for(let card of cards) {
                card.style.display = card.innerText.toUpperCase().includes(q) ? "block" : "none";
            }
        }
    </script>
</body>
</html>`;

fs.writeFileSync(path.join(outputDir, 'diccionario_excel.html'), html);
console.log(`Generado diccionario_excel.html con ${tableNames.length} tablas.`);
