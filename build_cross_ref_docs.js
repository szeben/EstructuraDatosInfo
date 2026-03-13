
const fs = require('fs');
const XLSX = require('xlsx');

// 1. ADVANCED RELATIONSHIP MAPPING (FK -> MASTER)
const relationshipMap = {
    'ID_EMPRESA': 'EO_EMPRESA', 'CIA_CODCIA': 'EO_EMPRESA', 'CIA_CODACT': 'EO_EMPRESA', 'CIAANT': 'EO_EMPRESA',
    'ID_PERSONA': 'EO_PERSONA', 'ID_PERSONA_HR': 'EO_PERSONA',
    'FICHA': 'TA_RELACION_LABORAL', 'TRAB_FICTRA': 'TA_RELACION_LABORAL', 'FICANT': 'TA_RELACION_LABORAL', 'FICMBA': 'TA_RELACION_LABORAL',
    'ID_TIPO_IDEN': 'EO_TIPO_IDENTIFICACION', 'ID_PAIS': 'SPI_PAISES', 'CODPAI': 'SPI_PAISES',
    'ID_LOCALIDAD': 'EO_LOCALIDAD', 'CODLOC': 'EO_LOCALIDAD',
    'ID_CARGO': 'EO_CARGO', 'CODCAR': 'EO_CARGO', 'CGO_CODCAR': 'EO_CARGO',
    'ID_UNIDAD': 'EO_UNIDAD', 'CODDEP': 'EO_UNIDAD', 'DPTO_CODDEP': 'EO_UNIDAD', 'CODUNI': 'EO_UNIDAD',
    'ID_NOMINA': 'NMT003', 'TNOM_TIPNOM': 'NMT003', 'TIPNOM': 'NMT003',
    'CTO_CODCTO': 'NMT027', 'CODCTO': 'NMT027', 'CTOMBA': 'NMT027',
    'ID_PROCESO': 'NMT011', 'PROC_TIPPRO': 'NMT011', 'TIPPRO': 'NMT011',
    'ID_PERFIL': 'SS_PERFIL', 'ID_USUARIO': 'SPI_USUARIO', 'USERID': 'SPI_USUARIO', 'USR_USERID': 'SPI_USUARIO',
    'ID_BANDA': 'RS_BANDA', 'ID_COMPETENCIA': 'RS_COMPETENCIA', 'ID_CV': 'RS_CV',
    'ID_REQUISICION': 'RS_REQUISICION', 'ID_RESUMEN': 'RS_RESUMEN',
    'ID_PARENTESCO': 'TA_PARENTESCOS', 'ID_PARIENTE': 'EO_PARIENTE',
    'ID_BENEFICIARIO': 'TA_BENEFICIARIOS', 'CODBEN': 'TA_BENEFICIARIOS',
    'CODE_PLAN': 'TA_PRIMA_PLAN', 'CODE_POLIZA': 'TA_POLIZAS', 'CODE_FINANCIA': 'TA_FINANCIAMIENTO_SEGURO'
};

// 2. LOAD EXCEL DATA
const workbook = XLSX.readFile('diccionario_datos_spi.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const allTables = {};
let currentT = '';
for (let i = 2; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    const tname = row[0] ? String(row[0]).trim() : '';
    if (tname) {
        currentT = tname;
        allTables[currentT] = { columns: [], referencedBy: [] };
    }
    const cname = row[2] ? String(row[2]).trim() : '';
    if (cname && currentT) {
        allTables[currentT].columns.push({
            name: cname,
            type: row[3] || '',
            len: row[4] || '',
            desc: row[6] || ''
        });
    }
}

// 3. BUILD INVERSE RELATIONSHIPS (REFERENCED BY)
for (const tname in allTables) {
    allTables[tname].columns.forEach(col => {
        const targetTable = relationshipMap[col.name];
        if (targetTable && allTables[targetTable] && targetTable !== tname) {
            if (!allTables[targetTable].referencedBy.includes(tname)) {
                allTables[targetTable].referencedBy.push(tname);
            }
        }
    });
}

// 4. GENERATE HTML
const sortedNames = Object.keys(allTables).sort();
let tableCards = "";

for (const tname of sortedNames) {
    const info = allTables[tname];
    
    let rows = info.columns.map(c => {
        const targetTable = relationshipMap[c.name];
        const isFk = targetTable && allTables[targetTable] && targetTable !== tname;
        const colDisplay = isFk ? `<a href="#${targetTable}" class="fk-link">${c.name}</a>` : c.name;
        const badge = isFk ? `<span class="fk-badge">FK &rarr; ${targetTable}</span>` : "";

        return `
            <tr>
                <td><span class="col-name">${colDisplay}</span> ${badge}</td>
                <td>${c.type}</td>
                <td>${c.len}</td>
                <td class="desc-cell">${c.desc}</td>
            </tr>
        `;
    }).join("");

    let refSection = "";
    if (info.referencedBy.length > 0) {
        refSection = `
        <div class="ref-by">
            <strong>Referenciada por:</strong>
            <div class="ref-links">
                ${info.referencedBy.sort().map(ref => `<a href="#${ref}">${ref}</a>`).join("")}
            </div>
        </div>
        `;
    }

    tableCards += `
    <div class="table-card" id="${tname}" data-name="${tname}">
        <div class="table-header-flex">
            <h3>AMCOR.${tname}</h3>
            <a href="#top" class="top-link">&uarr; Ir arriba</a>
        </div>
        <table class="data-table">
            <thead><tr><th>Campo & Relación</th><th>Tipo</th><th>Long</th><th>Descripción</th></tr></thead>
            <tbody>${rows}</tbody>
        </table>
        ${refSection}
    </div>
    `;
}

const css = `
:root {
    --primary: #6366f1; --bg-dark: #0f172a; --card-bg: #1e293b;
    --text-main: #f8fafc; --text-muted: #94a3b8; --accent: #38bdf8;
    --border: rgba(255, 255, 255, 0.1);
}
body { font-family: 'Inter', sans-serif; background: var(--bg-dark); color: var(--text-main); font-size: 0.7rem; margin: 0; scroll-behavior: smooth; }
.container { max-width: 1400px; margin: 0 auto; padding: 1rem; }
header { padding: 1rem; border-bottom: 2px solid var(--primary); text-align: center; }
.search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 0.5rem 0; z-index: 1000; border-bottom: 1px solid var(--border); }
.search-box { width: 100%; padding: 0.4rem; background: var(--card-bg); color: white; border: 1px solid var(--border); border-radius: 4px; }
.table-card { background: var(--card-bg); border-radius: 6px; padding: 0.8rem; margin-bottom: 2rem; border: 1px solid var(--border); transition: 0.3s; }
.table-card:target { border-color: var(--accent); box-shadow: 0 0 15px var(--accent); transform: scale(1.01); }
.data-table { width: 100%; border-collapse: collapse; }
.data-table th { text-align: left; padding: 0.4rem; background: rgba(99, 102, 241, 0.15); color: var(--accent); border-bottom: 1px solid var(--border); }
.data-table td { padding: 0.3rem; border-bottom: 1px solid var(--border); }
.col-name a { color: #facc15; text-decoration: none; border-bottom: 1px dotted #facc15; }
.fk-badge { font-size: 0.55rem; background: #10b981; color: white; padding: 1px 3px; border-radius: 3px; margin-left: 4px; }
.ref-by { margin-top: 1rem; padding: 0.5rem; background: rgba(0,0,0,0.2); border-radius: 4px; }
.ref-links { display: flex; flex-wrap: wrap; gap: 0.4rem; margin-top: 0.3rem; }
.ref-links a { font-size: 0.6rem; color: var(--accent); text-decoration: none; padding: 2px 4px; border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 3px; }
.ref-links a:hover { background: var(--accent); color: black; }
.top-link { font-size: 0.6rem; color: var(--accent); text-decoration: none; }
`;

const html = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Diccionario con Referencias Cruzadas - Infocent</title>
    <style>${css}</style>
</head>
<body>
    <div id="top"></div>
    <header><h1>DICCIONARIO CON REFERENCIAS CRUZADAS</h1></header>
    <div class="container">
        <nav style="display:flex; gap:10px; justify-content:center; margin-bottom:10px;">
            <a href="index.html" style="color:white; font-size:0.7rem;">&larr; Volver</a>
        </nav>
        <div class="search-container">
            <input type="text" id="masterSearch" class="search-box" placeholder="Buscar tabla o ver quién la referencia..." onkeyup="filter()">
        </div>
        <div id="tableList">${tableCards}</div>
    </div>
    <script>
        function filter() {
            let q = document.getElementById('masterSearch').value.toUpperCase();
            let cards = document.getElementsByClassName('table-card');
            for(let card of cards) { card.style.display = card.innerText.toUpperCase().includes(q) ? "block" : "none"; }
        }
    </script>
</body>
</html>`;

fs.writeFileSync('b:/Antigravity/BD_Infocent/docs/diccionario_excel.html', html);
console.log("Diccionario con referencias cruzadas generado correctamente.");
