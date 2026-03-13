
const fs = require('fs');
const XLSX = require('xlsx');

// 1. EXTRACT PK/FK FROM OCR TEXT
const ocrContent = `
AMCOR.NMT037
TXTR_PK (CIA_CODCIA, TNOM_TIPNOM, PROC_TIPPRO, FPRO_SUBPRO, FPRO_ANOCAL, FPRO_NUMPER, SUBTIP, FICTRA)
AMCOR.SPI_USUARIO
PK_USU (ID)
FK_USUARIO_PAIS (ID_PAIS) FK_USU_PERFIL (ID_PERFIL) FK_USU_PERSONA (ID_PERSONA)
AMCOR.TA_RELACION_LABORAL
LABORAL_PK (ID_EMPRESA, FICHA)
FK_HORARIO1_TRAB (ID_EMPRESA, HORARIO1) FK_HORARIO2_TRAB (ID_EMPRESA, HORARIO2) FK_HORARIO3_TRAB (ID_EMPRESA, HORARIO3) FK_HORARIO4_TRAB (ID_EMPRESA, HORARIO4) FK_HORARIO5_TRAB (ID_EMPRESA, HORARIO5) LABORAL_EMPRESA (ID_EMPRESA) LABORAL_GRADO (ID_EMPRESA, ID_GRADO) LABORAL_PERSONA (ID_PERSONA)
AMCOR.EO_CARGO
PK_CGO (ID_EMPRESA, ID)
FK_CARGO_CLASI (ID_EMPRESA, ID_CLASIFICA) FK_CARGO_EMPRE (ID_EMPRESA) FK_JERARQUIA_CARGO (ID_EMPRESA, ID_CARGO_SUP)
AMCOR.NMM019
MVAR_PK (CIA_CODCIA, TNOM_TIPNOM, PROC_TIPPRO, FPRO_SUBPRO, FPRO_ANOCAL, FPRO_NUMPER, USERMO, MODMOV, SECMOV, SUBSEC)
MVAR_CTO (CIA_CODCIA, TNOM_TIPNOM, CTO_CODCTO) MVAR_FPRO (CIA_CODCIA, TNOM_TIPNOM, PROC_TIPPRO, FPRO_SUBPRO, FPRO_ANOCAL, FPRO_NUMPER) MVAR_MOTHE (CIA_CODCIA, ID_MOTIVO_HE) MVAR_TABU (CIA_CODCIA, TABU_CODTAB) MVAR_TRAB_FICSUS (CIA_CODCIA, TRAB_FICSUS) MVAR_TRAB_FICTRA (CIA_CODCIA, TRAB_FICTRA)
AMCOR.EO_PERSONA
PK_PERSON (ID)
FK_PERSONA_TIPO_ID (ID_TIPO_IDEN)
AMCOR.EO_EMPRESA
PK_EMPRE (ID)
AMCOR.SS_ROL_OPCION
PK_ROL_OPC (ID_ROL, ID_SISTEMA, ID_MODULO, ID_OPCION)
FK_ROL_OPC_OPC (ID_SISTEMA, ID_MODULO, ID_OPCION) FK_ROL_OPC_ROL (ID_ROL)
AMCOR.SS_USUARIO_UNIDAD
PK_AUTO_UNI (ID_USUARIO, ID_EMPRESA, ID_UNIDAD)
FK_AUT_UNI_UNI (ID_EMPRESA, ID_UNIDAD) FK_AUT_UNI_USU (ID_USUARIO)
AMCOR.NM_RELACION_PAGO
REL_PAGO_PK (ID_EMPRESA, FICHA)
FK_CLASIF_DESCANSO (ID_EMPRESA, ID_DESCANSO)
AMCOR.NMT027
CTO_PK (CIA_CODCIA, TNOM_TIPNOM, CODCTO)
CTO_CTO_CTOAFE (CIA_CODCIA, TNOM_TIPNOM, CTOAFE) CTO_CTO_CTOAHO (CIA_CODCIA, TNOM_TIPNOM, CTOAHO) CTO_CTO_CTOFIJ (CIA_CODCIA, TNOM_TIPNOM, CTOFIJ) CTO_CTO_CTOINT (CIA_CODCIA, TNOM_TIPNOM, CTOINT) CTO_CTO_CTOPRO (CIA_CODCIA, TNOM_TIPNOM, CTOPRO) CTO_CTO_CTOSUP (CIA_CODCIA, TNOM_TIPNOM, CTOSUP)
AMCOR.TA_CONFIG_SEGUROS
PK_CFG_SEG (ID_EMPRESA, ID_EMPRE_SEGURO)
FK_CFG_SEG (ID_EMPRESA, ID_EMPRE_SEGURO)
AMCOR.TA_RAMOS_SEGUROS
PK_RAMOS_SEG (ID_EMPRESA, ID_EMPRE_SEGURO, ID)
FK_RAMOS_SEG (ID_EMPRESA, ID_EMPRE_SEGURO)
AMCOR.EO_UNIDAD
PK_UNI (ID_EMPRESA, ID)
FK_JERARQUIA_UNIDAD (ID_EMPRESA, ID_UNIDAD_SUP) FK_UNIDAD_EMPRE (ID_EMPRESA)
AMCOR.RS_HISTORIAL
PK_RS_HISTORIAL (ID)
FK_RS_HISTORIAL_REQ (ID_REQUISICION)
`;

function parseConstraints(text) {
    const tableConstraints = {};
    let currentTable = null;
    const lines = text.split('\n');
    for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        if (line.startsWith('AMCOR.')) {
            currentTable = line.replace('AMCOR.', '').trim();
            tableConstraints[currentTable] = { pks: [], fks: [] };
            continue;
        }
        if (!currentTable) continue;

        // PK patterns: PK_XXX (COLS) or XXX_PK (COLS)
        const pkMatch = line.match(/(PK_[A-Z0-9_]+|[A-Z0-9_]+_PK)\s*\(([^)]+)\)/);
        if (pkMatch) {
            const cols = pkMatch[2].split(',').map(s => s.trim());
            tableConstraints[currentTable].pks = [...new Set([...tableConstraints[currentTable].pks, ...cols])];
        }

        // FK patterns: FK_XXX (COLS) or REL_XXX (COLS)
        const fkMatches = line.matchAll(/(FK_[A-Z0-9_]+|[A-Z0-9_]+_FK)\s*\(([^)]+)\)/g);
        for (const m of fkMatches) {
            const cols = m[2].split(',').map(s => s.trim());
            tableConstraints[currentTable].fks = [...new Set([...tableConstraints[currentTable].fks, ...cols])];
        }
        
        // Also capture standalone FK calls like "LABORAL_EMPRESA (ID_EMPRESA)"
        const otherMatch = line.matchAll(/([A-Z0-9_]+)\s*\(([^)]+)\)/g);
        for (const m of otherMatch) {
            const label = m[1];
            const cols = m[2].split(',').map(s => s.trim());
            if (label.includes('PK')) continue; // handled
            if (label.includes('FK') || label.includes('REL') || label.includes('TRAB')) {
                 tableConstraints[currentTable].fks = [...new Set([...tableConstraints[currentTable].fks, ...cols])];
            }
        }
    }
    return tableConstraints;
}

const constraints = parseConstraints(ocrContent);

// 2. LOAD EXCEL DATA
const workbook = XLSX.readFile('diccionario_datos_spi.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const tables = {};
let currentTableName = '';

for (let i = 2; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    const maybeTableName = row[0] ? String(row[0]).trim() : '';
    const colName = row[2] ? String(row[2]).trim() : '';
    const colType = row[3] ? String(row[3]).trim() : '';
    const colLen = row[4];
    const colDesc = row[6] ? String(row[6]).trim() : '';

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
            desc: colDesc
        });
    }
}

// 3. GENERATE ENRICHED HTML
const tableNames = Object.keys(tables).sort();
let tableCards = "";

for (const tname of tableNames) {
    const info = tables[tname];
    const tableKey = tname.replace(/\s/g, ''); // OCR might not have spaces
    const tConstraints = constraints[tableKey] || constraints[tname] || { pks: [], fks: [] };
    
    let rows = info.columns.map(c => {
        const isPk = tConstraints.pks.includes(c.name);
        const isFk = tConstraints.fks.includes(c.name);
        const badge = (isPk ? '<span class="pk-badge">PK</span>' : '') + (isFk ? '<span class="fk-badge">FK</span>' : '');
        
        return `
            <tr class="${isPk ? 'pk-row' : ''}">
                <td><span class="col-name">${c.name}</span> ${badge}</td>
                <td>${c.type}</td>
                <td>${c.len || ''}</td>
                <td class="desc-cell">${c.desc || ''}</td>
            </tr>
        `;
    }).join("");

    tableCards += `
    <div class="table-card" id="${tname}" data-name="${tname}">
        <div class="table-header-flex">
            <h3>AMCOR.${tname}</h3>
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
body { font-family: 'Inter', sans-serif; background: var(--bg-dark); color: var(--text-main); font-size: 0.75rem; margin: 0; padding: 0; }
.container { max-width: 1400px; margin: 0 auto; padding: 1rem; }
header { padding: 1.5rem; border-bottom: 2px solid var(--primary); text-align: center; }
h1 { font-size: 1.5rem; color: var(--accent); margin:0; }
nav { display: flex; gap: 0.5rem; justify-content: center; padding: 0.5rem; }
nav a { color: white; text-decoration: none; padding: 0.2rem 0.6rem; border-radius: 4px; background: rgba(255,255,255,0.05); font-size: 0.7rem; }
.search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 0.5rem 0; z-index: 1000; border-bottom: 1px solid var(--border); }
.search-box { width: 100%; padding: 0.4rem; border-radius: 4px; border: 1px solid var(--border); background: var(--card-bg); color: white; }
.table-card { background: var(--card-bg); border-radius: 6px; padding: 0.8rem; margin-bottom: 2rem; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
.data-table { width: 100%; border-collapse: collapse; }
.data-table th { text-align: left; padding: 0.3rem; background: rgba(99, 102, 241, 0.2); color: var(--accent); border-bottom: 1px solid var(--border); font-weight: 800; text-transform: uppercase; font-size: 0.65rem; }
.data-table td { padding: 0.25rem; border-bottom: 1px solid var(--border); vertical-align: top; }
.col-name { color: #facc15; font-weight: 600; }
.pk-row { background: rgba(248, 113, 113, 0.05); }
.pk-badge { font-size: 0.6rem; background: #ef4444; color: white; padding: 1px 4px; border-radius: 3px; font-weight: bold; margin-left: 5px; }
.fk-badge { font-size: 0.6rem; background: #10b981; color: white; padding: 1px 4px; border-radius: 3px; font-weight: bold; margin-left: 5px; }
.desc-cell { color: var(--text-muted); line-height: 1.2; }
`;

const html = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Diccionario Enriquecido (PK/FK) - Infocent</title>
    <style>${css}</style>
</head>
<body>
    <header>
        <h1>DICCIONARIO TÉCNICO ENRIQUECIDO</h1>
        <p style="font-size: 0.7rem; color: var(--text-muted);">PK/FK integrados desde Modelo ER y Excel</p>
    </header>
    <div class="container">
        <nav><a href="index.html">Inicio</a><a href="relaciones.html">Relaciones</a></nav>
        <div class="search-container">
            <input type="text" id="masterSearch" class="search-box" placeholder="Ej: EO_PERSONA o NOMBRE1..." onkeyup="filter()">
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

fs.writeFileSync('b:/Antigravity/BD_Infocent/docs/diccionario_excel.html', html);
console.log("Generado diccionario_excel.html enriquecido.");
